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
  formula: 'dateAdd' | 'ifNotNull';
  // dateAdd
  sourceColumn?: string;
  addColumn?: string;
  unit?: 'day' | 'week' | 'month' | 'year';
  // ifNotNull: set to `then` column's value when `ifNotNull` column is non-empty, else `else` column
  ifNotNull?: string;
  then?: string;
  else?: string;
}

export interface CrossSheetLookup {
  /** Sheet whose column values to collect into a lookup set. */
  fromSheet: string;
  /** Column in fromSheet to collect (e.g. PatientId). */
  fromColumn: string;
  /** Column in this indicator's source sheet to check against the set. */
  matchColumn: string;
  /** Optional filters on fromSheet rows — only matching rows contribute to the lookup set. */
  lookupFilters?: FilterDef[];
}

export interface IndicatorDef {
  name: string;
  description?: string;
  source: string;
  requiredColumns: string[];
  computedColumns?: ComputedColumnDef[];
  filters: FilterDef[];
  anyOf?: FilterDef[];   // OR group: row counted if it passes at least one of these
  crossSheetLookup?: CrossSheetLookup;
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
  private categoryRanksConfig!: { categoryRanks: Record<string, string[]> };

  // Built from CombinedRADET during streaming; used to enrich HTS rows.
  // Persisted across runs via datim-cache.json so coverage grows over time.
  private datimToFacility = new Map<string, { Facility: string; State: string }>();
  private datimCachePath = '';

  // Maps the UUID suffix of compound values ("{facilityDATIM}_{uuid}") → facilityDATIM.
  // ExcelJS streaming sometimes returns just the UUID when its shared-string cache
  // misses the index; this persistent map lets us recover the DATIM in later runs.
  private compoundSuffixMap = new Map<string, string>();
  private suffixCachePath = '';

  // Set by preReadSharedStrings when xl/sharedStrings.xml could not be located
  // in a ZIP that clearly contains substantial worksheet data (i.e. this is not
  // a small inline-strings file where sharedStrings.xml legitimately doesn't
  // exist). Gates the end-of-run sanity check in process().
  private sharedStringsMissingButExpected = false;

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
    this.sheetHeaders = yaml.load(headersYaml) as Record<string, string[]>;
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

    this.suffixCachePath = path.join(this.options.configDir, 'compound-suffix-cache.json');
    this.loadSuffixCache();

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

    if (sheetName === 'CombinedRADET') {
      // LGA columns are excluded: they embed compound UIDs from OTHER facilities
      // (e.g. the LGA a patient transferred from) and would produce wrong matches.
      const skipCols = new Set(['LGA', 'LGAOfResidence']);

      // Check col5 (DATIMCode) first — it is the authoritative facility identifier.
      const col5Raw = String(record['DATIMCode'] ?? '').trim();
      const col5 = tryVal(col5Raw);
      if (col5) return col5;

      // Scan all remaining columns as fallbacks (handles ExcelJS shared-string
      // cache misses where col5 resolves to a non-DATIM value on a given run).
      for (const [col, val] of Object.entries(record)) {
        if (col === 'DATIMCode' || skipCols.has(col)) continue;
        const found = tryVal(String(val ?? '').trim());
        if (found) return found;
      }

      // Last resort for RADET: col5 might be a UUID that was previously seen as
      // a compound suffix (e.g. "facilityDATIM_patientUUID" → patientUUID).
      // Only try col5 to avoid false positives from shared option-concept UUIDs
      // that may appear in other columns.
      if (col5Raw && !col5Raw.includes('_')) {
        const suffixMapped = this.compoundSuffixMap.get(col5Raw);
        if (suffixMapped) return suffixMapped;
      }
    } else {
      for (const rawVal of Object.values(record)) {
        const found = tryVal(String(rawVal ?? '').trim());
        if (found) return found;
      }
      // Last resort for HTS: any column value might be a UUID that was seen
      // as a compound suffix in a previous run. All columns are checked because
      // HTS is the cross-sheet sheet with fewer ambiguity risks.
      for (const rawVal of Object.values(record)) {
        const str = String(rawVal ?? '').trim();
        if (!str) continue;
        const suffixMapped = this.compoundSuffixMap.get(str);
        if (suffixMapped) return suffixMapped;
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

  async process(onRow: (row: DashboardRow) => void): Promise<ProcessingStats> {
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

    // Group indicators by source sheet
    const bySheet = new Map<string, IndicatorDef[]>();
    for (const ind of this.indicatorsConfig.indicators) {
      if (!bySheet.has(ind.source)) bySheet.set(ind.source, []);
      bySheet.get(ind.source)!.push(ind);
    }

    // Determine all fields needed per indicator (groupBy + categorical filter cols)
    const indFilterCols = new Map<string, Set<string>>();
    for (const ind of this.indicatorsConfig.indicators) {
      const catFilters = new Set(
        ind.filters
          .filter(f => f.operator !== 'dateMode')
          .flatMap(f => [f.column, ...(f.ref ? [f.ref] : [])]),
      );
      // Also collect anyOf columns and their ref columns
      for (const f of (ind.anyOf ?? [])) {
        catFilters.add(f.column);
        if (f.ref) catFilters.add(f.ref);
      }
      // Track the matchColumn for cross-sheet lookup
      if (ind.crossSheetLookup) catFilters.add(ind.crossSheetLookup.matchColumn);
      indFilterCols.set(ind.name, catFilters);
    }

    // rawAccMap: indName → rawKeyStr → RawGroupEntry
    const rawAccMap = new Map<string, Map<string, RawGroupEntry>>();
    for (const ind of this.indicatorsConfig.indicators) rawAccMap.set(ind.name, new Map());

    // crossSheetSets: indName → Set of values collected from fromSheet for cross-sheet lookup
    const crossSheetSets = new Map<string, Set<string>>();
    for (const ind of this.indicatorsConfig.indicators) {
      if (ind.crossSheetLookup) crossSheetSets.set(ind.name, new Set());
    }

    // Org-unit raw data: col1-prefix → { stateRaw, lgaRaw, facilityRaw }
    const orgUnitSeen = new Map<string, { stateRaw: string; lgaRaw: string; facilityRaw: string }>();

    const seenSheets = new Set<string>();

    // Only keep in memory the rows needed by formula indicators as inputs.
    // All other rows are streamed immediately via onRow() to avoid OOM.
    const formulaInputIndicators = new Set<string>();
    for (const fd of this.indicatorsConfig.formulaIndicators ?? []) {
      formulaInputIndicators.add(fd.numerator);
      formulaInputIndicators.add(fd.denominator);
    }
    const formulaInputRows: DashboardRow[] = [];

    const buildDynamicMappings = (): Map<string, Map<string, string>> => {
      const dynamicMappings = new Map<string, Map<string, string>>();
      for (const [col, labels] of Object.entries(this.categoryRanksConfig.categoryRanks)) {
        const colFreq = freq.get(col);
        if (!colFreq || colFreq.size === 0) continue;
        const sorted = [...colFreq.entries()].sort((a, b) => b[1] - a[1]);
        const colMap = new Map<string, string>();
        for (let i = 0; i < Math.min(labels.length, sorted.length); i++) {
          const [normVal, count] = sorted[i];
          colMap.set(normVal, labels[i]);
          this.logger.info(
            `  [${col}] rank ${i + 1}: "${normVal}" (n=${count.toLocaleString()}) -> "${labels[i]}"`,
          );
        }
        dynamicMappings.set(col, colMap);
      }
      return dynamicMappings;
    };

    const finaliseIndicators = (
      indicatorsToFinalise: IndicatorDef[],
      dynamicMappings: Map<string, Map<string, string>>,
    ): void => {
      for (const ind of indicatorsToFinalise) {
        this.logger.info(`\n-- Finalising indicator: ${ind.name} --`);
        const indRaw = rawAccMap.get(ind.name)!;

        if (indRaw.size === 0) {
          this.logger.warn(`  No rows accumulated for ${ind.name} (all filtered by date range)`);
          continue;
        }

        // Categorical filters that apply to mapped values (exclude dateMode)
        const catFilters   = ind.filters.filter(f => f.operator !== 'dateMode');
        const anyOfFilters = ind.anyOf ?? [];

        // Final accumulator after mapping + re-grouping
        const finalAcc  = new Map<string, StreamAcc>();
        const finalComps = new Map<string, Record<string, string>>();

        for (const [, { rawComponents, acc }] of indRaw) {
          // Map each raw component to a human-readable label
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

          // Resolve org-unit names from facility DATIM code
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
            mapped['State']    = '';
            mapped['Facility'] = '';
            mapped['LGA']      = '';
            // Only use __rawDatimCode__ if it is a proper 11-char DATIM UID.
            const rawDatimCode = rawComponents['__rawDatimCode__'];
            const isValidDatim = !!rawDatimCode &&
              rawDatimCode.length === 11 &&
              /^[A-Za-z][A-Za-z0-9]{10}$/.test(rawDatimCode);
            mapped['DATIMCode'] = isValidDatim ? rawDatimCode : '';
            if (isValidDatim) {
              const rInfo = this.datimToFacility.get(rawDatimCode);
              if (rInfo) {
                mapped['Facility'] = rInfo.Facility;
                mapped['State']    = rInfo.State;
              }
            }
          }

          // Apply categorical filters on mapped values
          if (catFilters.length > 0 || anyOfFilters.length > 0) {
            const passedStr: Record<string, unknown> = {};
            for (const [k, v] of Object.entries(mapped)) passedStr[k] = v;
            if (catFilters.length > 0 && !this.filterEngine.passesAll(passedStr, catFilters)) continue;
            if (anyOfFilters.length > 0 && !this.filterEngine.passesAny(passedStr, anyOfFilters)) continue;
          }

          this.stats.filteredRows += acc.count;

          // Re-group by final (mapped) dimensions only
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

        let rowCount = 0;
        for (const [keyStr, acc] of finalAcc) {
          const kc    = finalComps.get(keyStr)!;
          const value = this.resolveAccValue(acc, ind);
          const row: DashboardRow = {
            Period:         kc['__period__'] ?? globalPeriod,
            State:          kc['State']      ?? '',
            Facility:       kc['Facility']   ?? '',
            DATIMCode:      kc['DATIMCode']  ?? '',
            Indicator:      ind.name,
            Disaggregation: ind.disaggregation ?? 'Total',
            Category:       kc[ind.disaggregation ?? ''] ?? 'Total',
            Sex:            kc['Sex']        ?? '',
            AgeBand:        kc['AgeBand']    ?? '',
            Value:          value,
            Numerator:      value,
            Denominator:    null,
            Target:         null,
            AchievementPct: null,
          };
          onRow(row);
          if (formulaInputIndicators.has(ind.name)) formulaInputRows.push(row);
          rowCount++;
        }

        this.stats.aggregatedRows += rowCount;
        this.stats.indicatorsProcessed.push(ind.name);
        this.logger.info(`  Generated ${rowCount} dashboard rows for ${ind.name}`);
      }
    };

    // ── SINGLE STREAMING PASS ─────────────────────────────────────────────────

    this.logger.info('AggregationEngine: single streaming pass (calibration + processing)…');

    // Pre-read the shared strings table from the raw ZIP before starting ExcelJS.
    // ACE-1_Combined_RADET files store all 5 worksheets BEFORE xl/sharedStrings.xml
    // in the ZIP, so ExcelJS sometimes only partially loads the 1.3M-entry table,
    // leaving high-index DATIM codes unresolved. We pre-read the full table and
    // inject it directly, bypassing ExcelJS's incremental parser.
    // Files without sharedStrings (e.g. plain RADET.xlsx with inline strings) are
    // handled gracefully — preloadedStrings is empty and the patch is skipped.
    // styles: 'ignore' prevents a crash on real RADET files where ExcelJS cannot
    // resolve styles before processing worksheets. Date values in RADET are stored
    // as text strings (dd/MM/yyyy) so style-based date detection is not required.
    //
    // xl/_rels/workbook.xml.rels is also stored AFTER the worksheet files in the ZIP,
    // so ExcelJS cannot map internal filenames (sheet1, sheet2…) to visible tab names
    // (CombinedRADET, CombinedHTS…) when it emits worksheet events. We pre-read both
    // workbook.xml and workbook.xml.rels to build our own filename→tabName map.
    const [{ strings: preloadedStrings, nsPrefix }, sheetNameMap] = await Promise.all([
      this.preReadSharedStrings(this.options.workbookPath),
      this.preReadSheetNames(this.options.workbookPath),
    ]);
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(
      this.options.workbookPath,
      { sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit' },
    );
    if (preloadedStrings.length > 0) {
      // Capture reader reference for closure (avoids 'this' binding issues in patched method)
      const readerRef = reader as unknown as Record<string, unknown>;
      const stringsRef = preloadedStrings;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (reader as any)._parseSharedStrings = async function*(entry: any) {
        readerRef['sharedStrings'] = stringsRef;
        // Drain the ZIP entry so the stream can advance to the next entry
        await new Promise<void>(resolve => {
          entry.on('end', () => resolve());
          entry.on('error', () => resolve());
          entry.resume();
        });
      };
    }
    if (nsPrefix) {
      // The worksheet XML uses the same namespace-prefixed style as
      // sharedStrings.xml (e.g. <x:row><x:c t="s"><x:v>0</x:v></x:c></x:row>).
      // ExcelJS's row/cell SAX handler only recognises unprefixed tag names,
      // so strip the prefix from each worksheet's raw XML stream before it
      // reaches the parser. See stripNamespacePrefix() for why this is safe.
      this.logger.info(`AggregationEngine: worksheet XML is namespace-prefixed ("${nsPrefix}") — rewriting tags before parsing`);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const readerAny = reader as any;
      const originalParseWorksheet = readerAny._parseWorksheet.bind(readerAny);
      const stripFn = this.stripNamespacePrefix.bind(this);
      readerAny._parseWorksheet = function* (iterator: AsyncIterable<Buffer | string>, sheetNo: string) {
        yield* originalParseWorksheet(stripFn(iterator, nsPrefix), sheetNo);
      };
    }
    reader.read();

    for await (const worksheetReader of reader) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rawSheetName: string = (worksheetReader as any).name;
      // Translate ExcelJS's internal filename (e.g. "sheet1") to the actual tab name
      // (e.g. "CombinedRADET"). Falls back to rawSheetName when the map is empty or
      // ExcelJS already resolved the name correctly.
      // ExcelJS generates "Sheet1", "Sheet2"… when it can't resolve tab names from
      // workbook.xml.rels. Our preReadSheetNames map uses lowercase keys ("sheet1"…)
      // matching the ZIP filenames. Normalise to lowercase for the lookup.
      const sheetName: string = sheetNameMap.get(rawSheetName.toLowerCase()) ?? rawSheetName;
      seenSheets.add(sheetName);

      const indicators = bySheet.get(sheetName);
      if (!indicators) {
        for await (const _ of worksheetReader) { /* drain */ }
        continue;
      }

      const predefined = this.sheetHeaders[sheetName] as string[] | undefined;
      const headers: string[] = predefined ? ['', ...predefined] : [];
      let isFirstRow = !predefined;

      this.logger.info(
        `  Streaming sheet: ${sheetName} (${indicators.map(i => i.name).join(', ')})` +
        (predefined ? ` [${predefined.length} positional cols]` : ' [first-row headers]'),
      );

      for await (const row of worksheetReader) {
        if (isFirstRow) {
          isFirstRow = false;
          (row as ExcelJS.Row).eachCell((cell, col) => {
            headers[col] = String(cell.value ?? '').trim();
          });
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

        // ── Apply computed columns for indicators on this sheet ──────────────
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

        // ── Build compound-suffix cache (Facility + DATIMCode columns only) ────────
        // Whenever ExcelJS correctly reads a compound "{facilityDATIM}_{uuid}" in a
        // facility-specific column, persist "uuid → facilityDATIM".  We restrict to
        // the Facility (col4) and DATIMCode (col5) columns only — other categorical
        // columns (Sex, ARTStatus, LGA, etc.) use SHARED option-concept UUIDs that
        // are identical across many facilities and would create false attributions.
        for (const col of ['Facility', 'DATIMCode'] as const) {
          const rawVal = record[col];
          const str = String(rawVal ?? '').trim();
          const datim = this.extractDatimFromCompound(str);
          if (datim && this.orgUnitHelper.lookupByDATIM(datim)) {
            const suffix = str.slice(12); // 11 DATIM chars + 1 underscore
            if (suffix && !this.compoundSuffixMap.has(suffix)) {
              this.compoundSuffixMap.set(suffix, datim);
            }
          }
        }

        // Populate RADET cross-sheet lookup (used for HTS facility info)
        if (sheetName === 'CombinedRADET' && rowDatim && !this.datimToFacility.has(rowDatim)) {
          const ou = this.orgUnitHelper.lookupByDATIM(rowDatim);
          this.datimToFacility.set(rowDatim, {
            Facility: ou?.facility ?? '',
            State:    ou?.state    ?? '',
          });
        }

        // Collect raw org-unit data for the CSV diagnostic output
        if (sheetName === 'CombinedRADET' && rowDatim && !orgUnitSeen.has(rowDatim)) {
          orgUnitSeen.set(rowDatim, {
            stateRaw:    String(record['State']    ?? '').trim(),
            lgaRaw:      String(record['LGA']      ?? '').trim(),
            facilityRaw: String(record['Facility'] ?? '').trim(),
          });
        }

        // ── Populate cross-sheet lookup sets ──────────────────────────────────
        for (const [indName, set] of crossSheetSets) {
          const lookup = this.indicatorsConfig.indicators.find(i => i.name === indName)!.crossSheetLookup!;
          if (lookup.fromSheet !== sheetName) continue;
          if (lookup.lookupFilters && lookup.lookupFilters.length > 0) {
            const normRecord: Record<string, unknown> = {};
            for (const [k, v] of Object.entries(record)) normRecord[k] = this.normalizeOptionValue(String(v ?? '').trim());
            if (!this.filterEngine.passesAll(normRecord, lookup.lookupFilters)) continue;
          }
          const val = String(record[lookup.fromColumn] ?? '').trim();
          if (val) set.add(val);
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

          // ── Cross-sheet lookup gate ───────────────────────────────────────────
          if (ind.crossSheetLookup) {
            const set = crossSheetSets.get(ind.name);
            const matchVal = String(record[ind.crossSheetLookup.matchColumn] ?? '').trim();
            if (!set || !matchVal || !set.has(matchVal)) continue;
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
            // Non-RADET sheets (e.g. HTS): DATIMCode column is the facility
            // identifier. Extract the 11-char DATIM prefix from compound values
            // (e.g. "JPBcTpp6XUu_uuid" → "JPBcTpp6XUu"). normalizeOptionValue
            // strips the prefix and keeps the UUID suffix — wrong for this field.
            __rawDatimCode__: sheetName !== 'CombinedRADET'
              ? (() => {
                  const raw = String(record['DATIMCode'] ?? '').trim();
                  return this.extractDatimFromCompound(raw) ?? raw;
                })()
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

      this.logger.info(`  Finished streaming ${sheetName}; finalising its indicators and releasing raw groups...`);
      const dynamicMappingsForSheet = buildDynamicMappings();
      finaliseIndicators(indicators, dynamicMappingsForSheet);
      for (const ind of indicators) rawAccMap.get(ind.name)!.clear();
    }


    // Log which orgUnits.yaml DATIMs were never seen in this run
    const allKnownDatims = this.orgUnitHelper.getAllDatimCodes();
    const neverSeen = allKnownDatims.filter(d => !orgUnitSeen.has(d));
    if (neverSeen.length > 0) {
      this.logger.warn(
        `  ${neverSeen.length} facilities from orgUnits.yaml had no rows resolved in this run ` +
        `(may have no data in this period, or ExcelJS cache miss for all their rows):`,
      );
      for (const d of neverSeen) {
        const ou = this.orgUnitHelper.lookupByDATIM(d);
        this.logger.warn(`    - ${ou?.facility ?? d} (${d})`);
      }
    } else {
      this.logger.info('  All 74 facilities resolved in this run.');
    }

    // Write diagnostic CSV so operators can verify / fill in missing org-unit names
    this.writeOrgUnitsRawCsv(orgUnitSeen);
    this.saveDatimCache();
    this.saveSuffixCache();

    // ── Post-process: map raw keys → labels, apply categorical filters ────────

    this.logger.info('AggregationEngine: post-processing accumulators…');

    for (const [sheetName] of bySheet) {
      if (!seenSheets.has(sheetName)) {
        const msg = `Worksheet '${sheetName}' not found in workbook`;
        this.logger.warn(msg);
        this.addValidationIssue(sheetName, 'N/A', 'Sheet', 'WARNING', msg);
      }
    }

    // ── Formula indicators ────────────────────────────────────────────────────
    if (this.indicatorsConfig.formulaIndicators?.length > 0) {
      this.logger.info('\n── Processing formula indicators ──');
      const formulaRows = this.formulaEngine.calculate(
        this.indicatorsConfig.formulaIndicators,
        formulaInputRows,
        globalPeriod,
      );
      for (const r of formulaRows) onRow(r);
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

    // ── Sanity gate ──────────────────────────────────────────────────────────
    // The authoritative signal that a run silently failed is 0 rows read from
    // a substantial workbook — that's grounds to abort here, before the
    // caller writes/uploads anything, rather than silently publishing an
    // empty or partial dataset that could overwrite good data.
    //
    // A missing/unparseable xl/sharedStrings.xml is NOT on its own proof of
    // failure: some export variants of this workbook (observed: whenever a
    // "CombineAHD" sheet is present) write every cell's text inline
    // (t="inlineStr") instead of via a shared-string table, which is a valid
    // OOXML format with no sharedStrings.xml entry at all — and ExcelJS reads
    // inline strings natively, independent of this class's shared-strings
    // pre-read. So it's only treated as fatal when it's corroborated by the
    // real symptom (0 rows actually read); otherwise it's logged and ignored.
    const workbookSizeBytes = fs.existsSync(this.options.workbookPath)
      ? fs.statSync(this.options.workbookPath).size
      : 0;
    if (this.stats.totalRows === 0 && workbookSizeBytes > AggregationEngine.SUBSTANTIAL_WORKSHEET_BYTES) {
      const sharedStringsNote = this.sharedStringsMissingButExpected
        ? ' xl/sharedStrings.xml could also not be read/parsed, which may be the cause.'
        : '';
      throw new Error(
        `AggregationEngine: aborting — 0 rows were read from a workbook of ` +
        `${workbookSizeBytes.toLocaleString()} bytes on disk. A file this size should contain data; ` +
        '0 rows most likely means something in the workbook could not be parsed (e.g. an XML style ' +
        'this pipeline does not recognise), not that today genuinely has no data.' + sharedStringsNote +
        ' Refusing to produce output from this run — investigate the workbook file directly.',
      );
    }
    if (this.sharedStringsMissingButExpected) {
      this.logger.warn(
        'AggregationEngine: xl/sharedStrings.xml could not be read/parsed, but ' +
        `${this.stats.totalRows.toLocaleString()} rows were read successfully — this workbook variant ` +
        'most likely uses inline strings instead of a shared-string table (a valid alternative OOXML ' +
        'format), not a parsing failure. Continuing.',
      );
    }

    return this.stats;
  }

  // ─── Private Helpers ─────────────────────────────────────────────────────

  /** Adds derived date fields to a record using dateAdd formula. */
  private applyComputedColumns(
    record: Record<string, unknown>,
    defs: ComputedColumnDef[],
  ): void {
    for (const def of defs) {
      if (def.formula === 'dateAdd') {
        const baseDate = this.dateHelper.parse(record[def.sourceColumn!]);
        if (!baseDate) continue;
        const addRaw = record[def.addColumn!];
        const addNum = typeof addRaw === 'number' ? addRaw : parseFloat(String(addRaw ?? ''));
        if (isNaN(addNum) || addNum <= 0) continue;
        record[def.name] = baseDate.add(addNum, def.unit!).toDate();
      } else if (def.formula === 'ifNotNull') {
        const checkVal = record[def.ifNotNull!];
        const hasValue = checkVal !== null && checkVal !== undefined && String(checkVal).trim() !== '';
        record[def.name] = hasValue ? record[def.then!] : record[def.else!];
      }
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

  // ─── Sheet-name pre-reader ────────────────────────────────────────────────
  // xl/_rels/workbook.xml.rels is stored AFTER the worksheet files in the ZIP.
  // ExcelJS streaming therefore cannot map internal filenames (sheet1…sheet5) to
  // visible tab names (CombinedRADET, CombinedHTS…) when it emits worksheet events.
  // We pre-read both workbook.xml and workbook.xml.rels to build that map ourselves.

  private async preReadSheetNames(xlsxPath: string): Promise<Map<string, string>> {
    // Returns: internal filename WITHOUT extension (e.g. "sheet1") → tab name (e.g. "CombinedRADET")
    const nameMap = new Map<string, string>();
    try {
      // eslint-disable-next-line @typescript-eslint/no-require-imports
      const unzipper = require('unzipper') as {
        Open: {
          file(p: string): Promise<{
            files: Array<{ path: string; buffer(): Promise<Buffer> }>;
          }>;
        };
      };
      const dir = await unzipper.Open.file(xlsxPath);

      const wbFile   = dir.files.find(f => f.path === 'xl/workbook.xml');
      const relsFile = dir.files.find(f => f.path === 'xl/_rels/workbook.xml.rels');
      if (!wbFile || !relsFile) {
        this.logger.warn('AggregationEngine: workbook.xml or workbook.xml.rels not found — sheet names may be wrong');
        return nameMap;
      }

      const [wbXml, relsXml] = await Promise.all([
        wbFile.buffer().then(b => b.toString('utf8')),
        relsFile.buffer().then(b => b.toString('utf8')),
      ]);

      // workbook.xml: build r:id → tab name
      const ridToName = new Map<string, string>();
      const sheetRe = /<[^>]*?sheet\s[^>]*/g;
      let m: RegExpExecArray | null;
      while ((m = sheetRe.exec(wbXml)) !== null) {
        const elem = m[0];
        const nameMatch = elem.match(/\bname="([^"]+)"/);
        const ridMatch  = elem.match(/\br:id="([^"]+)"/);
        if (nameMatch && ridMatch) ridToName.set(ridMatch[1], nameMatch[1]);
      }

      // workbook.xml.rels: r:id → internal filename (without extension)
      const relRe = /<Relationship\s[^>]*/g;
      while ((m = relRe.exec(relsXml)) !== null) {
        const elem = m[0];
        if (!elem.toLowerCase().includes('worksheet')) continue;
        const targetMatch = elem.match(/\bTarget="([^"]+)"/);
        const idMatch     = elem.match(/\bId="([^"]+)"/);
        if (!targetMatch || !idMatch) continue;
        const parts    = targetMatch[1].split('/');
        const filename = parts[parts.length - 1].replace('.xml', '');
        const tabName  = ridToName.get(idMatch[1]);
        if (tabName) nameMap.set(filename, tabName);
      }

      this.logger.info(
        `AggregationEngine: pre-loaded ${nameMap.size} sheet name mappings ` +
        `(${[...nameMap.entries()].map(([k, v]) => `${k}→${v}`).join(', ')})`,
      );
    } catch (err) {
      this.logger.warn(`AggregationEngine: failed to pre-read sheet names (${String(err)})`);
    }
    return nameMap;
  }

  // ─── Shared-string pre-reader ────────────────────────────────────────────
  // Reads xl/sharedStrings.xml directly from the XLSX ZIP using unzipper (already
  // a transitive dependency of ExcelJS).  Returns a plain string array whose indices
  // match the shared-string references in the worksheet XML.

  // A worksheet entry above this size indicates a real, data-filled sheet
  // (as opposed to a small/empty sample or inline-strings-only workbook) —
  // used to decide whether a missing sharedStrings.xml is suspicious or benign.
  private static readonly SUBSTANTIAL_WORKSHEET_BYTES = 500_000;
  private static readonly SHARED_STRINGS_MAX_ATTEMPTS = 3;
  private static readonly SHARED_STRINGS_RETRY_DELAY_MS = 750;

  private async preReadSharedStrings(xlsxPath: string): Promise<{ strings: string[]; nsPrefix: string }> {
    this.logger.info('AggregationEngine: pre-reading xl/sharedStrings.xml from ZIP…');
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const unzipper = require('unzipper') as {
      Open: {
        file(p: string): Promise<{
          files: Array<{ path: string; uncompressedSize?: number; buffer(): Promise<Buffer> }>;
        }>;
      };
    };

    for (let attempt = 1; attempt <= AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS; attempt++) {
      try {
        const dir = await unzipper.Open.file(xlsxPath);
        const ssFile = dir.files.find(f => f.path === 'xl/sharedStrings.xml');

        if (!ssFile) {
          const hasSubstantialWorksheet = dir.files.some(
            f => f.path.startsWith('xl/worksheets/') &&
              (f.uncompressedSize ?? 0) > AggregationEngine.SUBSTANTIAL_WORKSHEET_BYTES,
          );

          if (attempt < AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS) {
            this.logger.warn(
              `AggregationEngine: xl/sharedStrings.xml not found in ZIP (attempt ${attempt}/${AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS}) — retrying…`,
            );
            await new Promise(resolve => setTimeout(resolve, AggregationEngine.SHARED_STRINGS_RETRY_DELAY_MS));
            continue;
          }

          if (hasSubstantialWorksheet) {
            this.sharedStringsMissingButExpected = true;
            this.logger.error(
              'AggregationEngine: xl/sharedStrings.xml not found after ' +
              `${AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS} attempts, but this workbook has ` +
              'substantial worksheet data — shared-string text (facility names, statuses, etc.) ' +
              'will be unreadable. This will be treated as a failed run rather than empty data.',
            );
          } else {
            this.logger.warn('AggregationEngine: xl/sharedStrings.xml not found in ZIP (no substantial worksheet data — likely inline-strings file)');
          }
          return { strings: [], nsPrefix: '' };
        }

        const xml = (await ssFile.buffer()).toString('utf8');

        // Some export tools (seen on this project: a "Combined RADET" merge
        // step) emit every element with an explicit namespace prefix, e.g.
        // <x:sst xmlns:x="..."><x:si><x:t>...</x:t></x:si></x:sst> instead of
        // the more common unprefixed <sst><si><t>...</t></si></sst>. Detect
        // whichever style this file uses from its root element so the <si>/<t>
        // boundary searches below match either form.
        const prefixMatch = xml.match(/<([A-Za-z0-9_]+):sst[\s>]/);
        const nsPrefix = prefixMatch ? `${prefixMatch[1]}:` : '';
        const siOpenTag  = `<${nsPrefix}si`;
        const siCloseTag = `</${nsPrefix}si>`;
        const tRe = new RegExp(`<${nsPrefix}t[^>]*>([^<]*)</${nsPrefix}t>`, 'g');

        const strings: string[] = [];
        let pos = 0;

        while (pos < xml.length) {
          const siStart = xml.indexOf(siOpenTag, pos);
          if (siStart < 0) break;
          const siEnd = xml.indexOf(siCloseTag, siStart);
          if (siEnd < 0) break;

          // Extract all <t>…</t> text runs within this <si> element (handles both
          // plain-text and rich-text formats; rich-text runs are concatenated).
          const siContent = xml.substring(siStart, siEnd);
          let text = '';
          tRe.lastIndex = 0;
          let m: RegExpExecArray | null;
          while ((m = tRe.exec(siContent)) !== null) {
            text += m[1]
              .replace(/&amp;/g, '&')
              .replace(/&lt;/g, '<')
              .replace(/&gt;/g, '>')
              .replace(/&apos;/g, "'")
              .replace(/&quot;/g, '"');
          }
          strings.push(text);
          pos = siEnd + siCloseTag.length;
        }

        this.logger.info(
          `AggregationEngine: pre-loaded ${strings.length.toLocaleString()} shared strings` +
          (nsPrefix ? ` (namespace-prefixed XML: "${nsPrefix}")` : '') +
          (attempt > 1 ? ` (succeeded on attempt ${attempt})` : ''),
        );

        // Same failure signature as "entry not found": the ZIP entry exists and
        // has real content, but our extraction found nothing usable. Without
        // this check a parsing regression (e.g. an unrecognised XML style)
        // would silently degrade to the same "0 rows, upload anyway" outcome
        // this gate exists to prevent.
        if (strings.length === 0 && (ssFile.uncompressedSize ?? 0) > AggregationEngine.SUBSTANTIAL_WORKSHEET_BYTES) {
          this.sharedStringsMissingButExpected = true;
          this.logger.error(
            `AggregationEngine: xl/sharedStrings.xml is ${(ssFile.uncompressedSize ?? 0).toLocaleString()} bytes ` +
            'but 0 strings were extracted from it — the XML style is likely not recognised by this parser. ' +
            'This will be treated as a failed run rather than empty data.',
          );
        }

        return { strings, nsPrefix };
      } catch (err) {
        if (attempt < AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS) {
          this.logger.warn(
            `AggregationEngine: failed to pre-read shared strings on attempt ${attempt}/${AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS} (${String(err)}) — retrying…`,
          );
          await new Promise(resolve => setTimeout(resolve, AggregationEngine.SHARED_STRINGS_RETRY_DELAY_MS));
          continue;
        }
        this.logger.warn(
          `AggregationEngine: failed to pre-read shared strings after ${AggregationEngine.SHARED_STRINGS_MAX_ATTEMPTS} attempts (${String(err)}); falling back to ExcelJS default`,
        );
        return { strings: [], nsPrefix: '' };
      }
    }
    return { strings: [], nsPrefix: '' };
  }

  // ─── Namespace-prefixed worksheet XML transform ─────────────────────────
  // Some export tools emit every OOXML element with an explicit namespace
  // prefix (e.g. <x:row><x:c t="s"><x:v>0</x:v></x:c></x:row>). ExcelJS's SAX
  // parser (saxes, non-namespace-aware) treats "x:row" as the literal tag
  // name, which never matches its hardcoded 'row'/'c'/'v' cases — so every
  // row is silently skipped. Since a raw '<' can never appear in XML text
  // content (it must be escaped as "&lt;"), stripping "<x:" -> "<" and
  // "</x:" -> "</" as bytes stream through is always safe: every match is
  // guaranteed to be an actual tag boundary, never a false positive in text.
  // Longest suffix of `text` that is a proper prefix of `pattern` (length
  // 1..pattern.length-1) — i.e. how many trailing characters of `text` could
  // be the start of an as-yet-incomplete match of `pattern`. Returns 0 if
  // there's no such overlap, meaning it's safe to emit all of `text`.
  //
  // A fixed holdback of "pattern.length - 1" characters is NOT sufficient:
  // if `text` happens to end with the pattern occurring COMPLETELY (e.g.
  // "...</x:"), a fixed holdback would still chop off the pattern's leading
  // "<" into the emitted portion and the rest into carry, so neither this
  // chunk's replace (missing the tail) nor the next chunk's replace
  // (missing the now-detached leading "<") ever sees the full pattern —
  // silently leaving it unreplaced in the output. Checking for an actual
  // partial-match overlap (rather than blindly holding back a fixed count)
  // avoids this: a complete match at the tail has zero overlap and is
  // replaced immediately; only a genuinely incomplete tail gets held back.
  private static longestPrefixOverlap(text: string, pattern: string): number {
    const maxLen = Math.min(pattern.length - 1, text.length);
    for (let len = maxLen; len >= 1; len--) {
      if (text.slice(text.length - len) === pattern.slice(0, len)) return len;
    }
    return 0;
  }

  private async *stripNamespacePrefix(
    iterator: AsyncIterable<Buffer | string>,
    nsPrefix: string,
  ): AsyncGenerator<string> {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { StringDecoder } = require('string_decoder') as { StringDecoder: new (enc: string) => { write(b: Buffer): string; end(): string } };
    const decoder = new StringDecoder('utf8');
    const openTag = `<${nsPrefix}`;
    const closeTag = `</${nsPrefix}`;
    let carry = '';

    for await (const chunk of iterator) {
      const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      const text = carry + decoder.write(buf);
      const holdback = Math.max(
        AggregationEngine.longestPrefixOverlap(text, openTag),
        AggregationEngine.longestPrefixOverlap(text, closeTag),
      );
      const safeEnd = text.length - holdback;
      if (safeEnd <= 0) {
        carry = text;
        continue;
      }
      const toEmit = text.slice(0, safeEnd);
      carry = text.slice(safeEnd);
      yield toEmit.split(openTag).join('<').split(closeTag).join('</');
    }
    const rest = carry + decoder.end();
    if (rest) {
      yield rest.split(openTag).join('<').split(closeTag).join('</');
    }
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

  private loadSuffixCache(): void {
    try {
      const raw = fs.readFileSync(this.suffixCachePath, 'utf8');
      const cache = JSON.parse(raw) as Record<string, string>;
      let loaded = 0;
      for (const [suffix, datim] of Object.entries(cache)) {
        if (!this.compoundSuffixMap.has(suffix)) {
          this.compoundSuffixMap.set(suffix, datim);
          loaded++;
        }
      }
      if (loaded > 0) {
        this.logger.info(
          `AggregationEngine: loaded ${loaded} compound-suffix entries from cache (${Object.keys(cache).length} total)`,
        );
      }
    } catch {
      // Cache not yet created — first run
    }
  }

  private saveSuffixCache(): void {
    let existing: Record<string, string> = {};
    try {
      existing = JSON.parse(fs.readFileSync(this.suffixCachePath, 'utf8'));
    } catch { /* new file */ }

    let added = 0;
    for (const [suffix, datim] of this.compoundSuffixMap.entries()) {
      if (!existing[suffix]) {
        existing[suffix] = datim;
        added++;
      }
    }

    fs.writeFileSync(this.suffixCachePath, JSON.stringify(existing, null, 2), 'utf8');
    if (added > 0) {
      this.logger.info(
        `AggregationEngine: saved ${added} new compound-suffix entries (${Object.keys(existing).length} total)`,
      );
    }
  }
}
