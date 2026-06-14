import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';

export interface AgeBandDef {
  label: string;
  min: number;
  max: number;
}

export interface AgeBandsConfig {
  ageBands: AgeBandDef[];
  defaultBand: string;
}

export class AgeBandHelper {
  private bands: AgeBandDef[];
  private defaultBand: string;

  constructor(configPath?: string) {
    const cfgPath = configPath ?? path.resolve(__dirname, '../config/ageBands.yaml');
    const raw = fs.readFileSync(cfgPath, 'utf8');
    const cfg = yaml.load(raw) as AgeBandsConfig;
    this.bands = cfg.ageBands;
    this.defaultBand = cfg.defaultBand ?? 'Unknown';
  }

  /** Assign PEPFAR age band from a raw age value (numeric or string) */
  getBand(rawAge: unknown): string {
    if (rawAge === null || rawAge === undefined || rawAge === '') return this.defaultBand;

    const str = String(rawAge).trim();

    // Some systems store infants as the string "<1", "< 1", "<01", or "0-1"
    if (/^<\s*0*1$/.test(str) || str === '0-1') return '<1';

    const age = typeof rawAge === 'number' ? rawAge : parseFloat(str);
    if (isNaN(age) || age < 0) return this.defaultBand;

    const floorAge = Math.floor(age);

    for (const band of this.bands) {
      if (floorAge >= band.min && floorAge <= band.max) {
        return band.label;
      }
    }

    return this.defaultBand;
  }

  /** Return all defined band labels in order */
  getAllBands(): string[] {
    return this.bands.map((b) => b.label);
  }

  /** Check if a given label is a valid age band */
  isValid(label: string): boolean {
    return this.bands.some((b) => b.label === label) || label === this.defaultBand;
  }
}
