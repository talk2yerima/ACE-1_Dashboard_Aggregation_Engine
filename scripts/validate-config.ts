import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';

type Severity = 'ERROR' | 'WARNING';

interface ValidationMessage {
  severity: Severity;
  location: string;
  message: string;
}

interface FilterDef {
  column?: string;
  ref?: string;
  valueColumn?: string;
}

interface ComputedColumnDef {
  name?: string;
  sourceColumn?: string;
  addColumn?: string;
  ifNotNull?: string;
  then?: string;
  else?: string;
}

interface CrossSheetLookupDef {
  fromSheet?: string;
  fromColumn?: string;
  matchColumn?: string;
  lookupFilters?: FilterDef[];
}

interface IndicatorDef {
  name?: string;
  source?: string;
  requiredColumns?: string[];
  computedColumns?: ComputedColumnDef[];
  filters?: FilterDef[];
  anyOf?: FilterDef[];
  crossSheetLookup?: CrossSheetLookupDef;
  groupBy?: string[];
  aggregation?: string;
  disaggregation?: string;
  valueColumn?: string;
  periodColumn?: string;
}

interface FormulaIndicatorDef {
  name?: string;
  numerator?: string;
  denominator?: string;
  formula?: string;
  outputType?: string;
  groupBy?: string[];
}

interface IndicatorsConfig {
  indicators?: IndicatorDef[];
  formulaIndicators?: FormulaIndicatorDef[];
}

const configDir = path.resolve(process.cwd(), 'config');
const derivedColumns = new Set(['AgeBand', 'Facility', 'State', 'LGA', 'DATIMCode']);
const validAggregations = new Set(['COUNT', 'SUM', 'AVG', 'MIN', 'MAX', 'COUNTDISTINCT']);
const validFormulaOutputTypes = new Set(['percentage', 'count', 'ratio']);
const validFormulaExpressions = new Set([
  'numerator / denominator',
  'numerator - denominator',
  'numerator + denominator',
  'numerator * denominator',
]);

function loadYaml<T>(fileName: string): T {
  const filePath = path.join(configDir, fileName);
  return yaml.load(fs.readFileSync(filePath, 'utf8')) as T;
}

function add(messages: ValidationMessage[], severity: Severity, location: string, message: string): void {
  messages.push({ severity, location, message });
}

function assertColumn(
  messages: ValidationMessage[],
  sheetHeaders: Map<string, Set<string>>,
  source: string,
  computed: Set<string>,
  location: string,
  column: string | undefined,
): void {
  if (!column) return;
  const headers = sheetHeaders.get(source);
  if (!headers) return;
  if (headers.has(column) || computed.has(column) || derivedColumns.has(column)) return;
  add(messages, 'ERROR', location, `Unknown column '${column}' for source '${source}'`);
}

function validateFilters(
  messages: ValidationMessage[],
  sheetHeaders: Map<string, Set<string>>,
  source: string,
  computed: Set<string>,
  location: string,
  filters: FilterDef[] | undefined,
): void {
  for (const filter of filters ?? []) {
    assertColumn(messages, sheetHeaders, source, computed, location, filter.column);
    assertColumn(messages, sheetHeaders, source, computed, location, filter.ref);
    assertColumn(messages, sheetHeaders, source, computed, location, filter.valueColumn);
  }
}

function main(): void {
  const messages: ValidationMessage[] = [];
  const indicatorsConfig = loadYaml<IndicatorsConfig>('indicators.yaml');
  const rawHeaders = loadYaml<Record<string, string[]>>('sheetHeaders.yaml');
  const sheetHeaders = new Map(
    Object.entries(rawHeaders).map(([sheet, headers]) => [sheet, new Set(headers)]),
  );

  const indicators = indicatorsConfig.indicators ?? [];
  const formulas = indicatorsConfig.formulaIndicators ?? [];
  const indicatorNames = new Set<string>();

  for (const [idx, indicator] of indicators.entries()) {
    const location = `indicators[${idx}]${indicator.name ? ` ${indicator.name}` : ''}`;
    if (!indicator.name) {
      add(messages, 'ERROR', location, 'Missing indicator name');
      continue;
    }
    if (indicatorNames.has(indicator.name)) {
      add(messages, 'ERROR', location, `Duplicate indicator name '${indicator.name}'`);
    }
    indicatorNames.add(indicator.name);

    if (!indicator.source) {
      add(messages, 'ERROR', location, 'Missing source sheet');
      continue;
    }
    if (!sheetHeaders.has(indicator.source)) {
      add(messages, 'ERROR', location, `Unknown source sheet '${indicator.source}'`);
      continue;
    }

    if (!indicator.aggregation || !validAggregations.has(indicator.aggregation)) {
      add(messages, 'ERROR', location, `Invalid aggregation '${indicator.aggregation ?? ''}'`);
    }

    const computed = new Set((indicator.computedColumns ?? []).map((col) => col.name).filter(Boolean) as string[]);
    for (const column of indicator.requiredColumns ?? []) {
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, column);
    }
    for (const column of indicator.groupBy ?? []) {
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, column);
    }
    assertColumn(messages, sheetHeaders, indicator.source, computed, location, indicator.disaggregation);
    assertColumn(messages, sheetHeaders, indicator.source, computed, location, indicator.valueColumn);
    assertColumn(messages, sheetHeaders, indicator.source, computed, location, indicator.periodColumn);

    for (const computedColumn of indicator.computedColumns ?? []) {
      if (!computedColumn.name) add(messages, 'ERROR', location, 'Computed column is missing name');
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, computedColumn.sourceColumn);
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, computedColumn.addColumn);
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, computedColumn.ifNotNull);
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, computedColumn.then);
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, computedColumn.else);
    }

    validateFilters(messages, sheetHeaders, indicator.source, computed, location, indicator.filters);
    validateFilters(messages, sheetHeaders, indicator.source, computed, location, indicator.anyOf);

    if (indicator.crossSheetLookup) {
      const lookup = indicator.crossSheetLookup;
      if (!lookup.fromSheet || !sheetHeaders.has(lookup.fromSheet)) {
        add(messages, 'ERROR', location, `Unknown cross-sheet lookup source '${lookup.fromSheet ?? ''}'`);
      } else {
        assertColumn(messages, sheetHeaders, lookup.fromSheet, new Set(), location, lookup.fromColumn);
        validateFilters(messages, sheetHeaders, lookup.fromSheet, new Set(), location, lookup.lookupFilters);
      }
      assertColumn(messages, sheetHeaders, indicator.source, computed, location, lookup.matchColumn);
    }
  }

  for (const [idx, formula] of formulas.entries()) {
    const location = `formulaIndicators[${idx}]${formula.name ? ` ${formula.name}` : ''}`;
    if (!formula.name) add(messages, 'ERROR', location, 'Missing formula indicator name');
    if (!formula.numerator) add(messages, 'ERROR', location, 'Missing numerator');
    if (!formula.denominator) add(messages, 'ERROR', location, 'Missing denominator');
    if (formula.numerator && !indicatorNames.has(formula.numerator)) {
      add(messages, 'ERROR', location, `Unknown numerator indicator '${formula.numerator}'`);
    }
    if (formula.denominator && !indicatorNames.has(formula.denominator)) {
      add(messages, 'ERROR', location, `Unknown denominator indicator '${formula.denominator}'`);
    }
    if (!formula.formula || !validFormulaExpressions.has(formula.formula)) {
      add(messages, 'ERROR', location, `Invalid formula expression '${formula.formula ?? ''}'`);
    }
    if (!formula.outputType || !validFormulaOutputTypes.has(formula.outputType)) {
      add(messages, 'ERROR', location, `Invalid outputType '${formula.outputType ?? ''}'`);
    }
    for (const field of formula.groupBy ?? []) {
      if (!['Period', 'State', 'Facility', 'DATIMCode', 'Indicator', 'Disaggregation', 'Category', 'Sex', 'AgeBand'].includes(field)) {
        add(messages, 'WARNING', location, `Formula groupBy '${field}' is not a standard dashboard output column`);
      }
    }
  }

  if (messages.length === 0) {
    console.log(`Config validation passed (${indicators.length} indicators, ${formulas.length} formulas).`);
    return;
  }

  for (const message of messages) {
    console.log(`[${message.severity}] ${message.location}: ${message.message}`);
  }

  const errors = messages.filter((message) => message.severity === 'ERROR').length;
  if (errors > 0) {
    console.error(`Config validation failed with ${errors} error(s).`);
    process.exit(1);
  }
  console.log(`Config validation passed with ${messages.length} warning(s).`);
}

main();
