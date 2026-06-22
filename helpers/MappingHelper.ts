import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';

interface MappingsConfig {
  mappings: Record<string, Record<string, string>>;
  dateFormats?: string[];
}

export class MappingHelper {
  private mappings: Record<string, Record<string, string>>;
  public dateFormats: string[];

  constructor(configPath?: string) {
    const cfgPath = configPath ?? path.resolve(__dirname, '../config/mappings.yaml');
    const raw = fs.readFileSync(cfgPath, 'utf8');
    const cfg = yaml.load(raw) as MappingsConfig;
    this.mappings = cfg.mappings ?? {};
    this.dateFormats = cfg.dateFormats ?? ['YYYY-MM-DD'];
  }

  /** Map a raw column value to its standardized form */
  map(column: string, rawValue: unknown): string {
    if (rawValue === null || rawValue === undefined) return '';
    const str = String(rawValue).trim();
    const colMap = this.mappings[column];
    if (!colMap) return str;
    return colMap[str] ?? str;
  }

  /** Check if a mapping exists for a column */
  hasMapping(column: string): boolean {
    return !!this.mappings[column];
  }

  /** Return all mappings for a column */
  getColumnMappings(column: string): Record<string, string> | undefined {
    return this.mappings[column];
  }

  /** Add or override a mapping at runtime */
  setMapping(column: string, rawValue: string, mappedValue: string): void {
    if (!this.mappings[column]) {
      this.mappings[column] = {};
    }
    this.mappings[column][rawValue] = mappedValue;
  }
}
