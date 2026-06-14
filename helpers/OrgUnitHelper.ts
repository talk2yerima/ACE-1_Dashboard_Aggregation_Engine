import fs from 'fs';
import yaml from 'js-yaml';

interface DATIMEntry {
  facility?: string;
  state?: string;
  lga?: string;
}

interface OrgUnitConfig {
  orgUnits?: {
    datimCodes?: Record<string, DATIMEntry>;
    states?: Record<string, string>;
    lgas?: Record<string, string>;
    facilities?: Record<string, string>;
  };
}

/**
 * Maps DHIS2 org-unit UIDs and DATIM facility codes to human-readable names.
 * Loaded from config/orgUnits.yaml.
 *
 * Priority (highest → lowest):
 *   1. datimCodes lookup  — keyed by the 11-char DATIM facility code
 *   2. uid-based maps     — states / lgas / facilities sections
 *   3. raw value          — fall-through when nothing matches
 */
export class OrgUnitHelper {
  private datimCodes  = new Map<string, DATIMEntry>();
  private states      = new Map<string, string>();
  private lgas        = new Map<string, string>();
  private facilities  = new Map<string, string>();

  constructor(filePath: string) {
    if (!fs.existsSync(filePath)) return;

    const cfg = yaml.load(fs.readFileSync(filePath, 'utf8')) as OrgUnitConfig;
    const ou = cfg?.orgUnits;
    if (!ou) return;

    for (const [code, entry] of Object.entries(ou.datimCodes ?? {})) {
      this.datimCodes.set(code.trim(), entry);
    }
    for (const [uid, name] of Object.entries(ou.states    ?? {})) this.states.set(uid.trim(),     name);
    for (const [uid, name] of Object.entries(ou.lgas      ?? {})) this.lgas.set(uid.trim(),       name);
    for (const [uid, name] of Object.entries(ou.facilities ?? {})) this.facilities.set(uid.trim(), name);
  }

  /** Resolve Facility / State / LGA via the DATIM facility code (most reliable). */
  lookupByDATIM(code: string): DATIMEntry | null {
    return this.datimCodes.get(code) ?? null;
  }

  mapState(raw: string):    string { return this.states.get(raw)     ?? raw; }
  mapLGA(raw: string):      string { return this.lgas.get(raw)       ?? raw; }
  mapFacility(raw: string): string { return this.facilities.get(raw) ?? raw; }

  get hasAnyMappings(): boolean {
    return (
      this.datimCodes.size > 0 ||
      this.states.size > 0 ||
      this.lgas.size > 0 ||
      this.facilities.size > 0
    );
  }
}
