import fs from 'fs';
import { Pool, PoolConfig } from 'pg';
import dayjs from 'dayjs';
import type { Logger } from 'winston';
import { DashboardRow } from './AggregationEngine';

interface EdctDbRow {
  Section:       string;
  State:         string;
  Facility:      string;
  Datim:         string;
  Sex:           string;
  AgeGroup:      string;
  Value:         string | number;
  DataElementId: string;
  ReportingDate: Date | string;
}

/**
 * Parses both .NET ADO.NET format and standard pg URL/key=value formats.
 *
 * .NET format:  Host=localhost;Port=5432;Database=baycentral;Username=dbadmin;Password=secret
 * pg URL:       postgresql://dbadmin:secret@localhost:5432/baycentral
 * pg key=value: host=localhost port=5432 dbname=baycentral user=dbadmin password=secret
 */
function parseConnectionString(connStr: string): PoolConfig {
  const trimmed = connStr.trim();

  // Already a URL
  if (trimmed.startsWith('postgres://') || trimmed.startsWith('postgresql://')) {
    return { connectionString: trimmed };
  }

  // Detect .NET ADO.NET style (semicolon-separated Key=Value)
  if (trimmed.includes(';') || /^[A-Z][a-zA-Z]+=/.test(trimmed)) {
    const cfg: PoolConfig = {};
    for (const part of trimmed.split(';')) {
      const eq = part.indexOf('=');
      if (eq < 0) continue;
      const key = part.slice(0, eq).trim().toLowerCase();
      const val = part.slice(eq + 1).trim();
      if (!val) continue;
      switch (key) {
        case 'host':     case 'server':   cfg.host     = val; break;
        case 'port':                      cfg.port     = Number(val); break;
        case 'database': case 'db':       cfg.database = val; break;
        case 'username': case 'user id':
        case 'user':     case 'userid':   cfg.user     = val; break;
        case 'password': case 'pwd':      cfg.password = val; break;
        // Silently ignore ADO.NET-only keys like "Include Error Detail"
      }
    }
    return cfg;
  }

  // Assume libpq key=value style
  return { connectionString: trimmed };
}

export class EdctService {
  private readonly config: PoolConfig;
  private readonly logger: Logger;

  constructor(connectionString: string, logger: Logger) {
    this.config = parseConnectionString(connectionString);
    this.logger = logger;
  }

  /** Returns true if the database is reachable, false otherwise. */
  async testConnection(): Promise<boolean> {
    const pool = new Pool({ ...this.config, connectionTimeoutMillis: 5000 });
    try {
      const client = await pool.connect();
      await client.query('SELECT 1');
      client.release();
      this.logger.info('  PostgreSQL connection test: OK');
      return true;
    } catch (err) {
      this.logger.warn(`  PostgreSQL connection test failed: ${String(err)}`);
      return false;
    } finally {
      await pool.end();
    }
  }

  async fetchDashboardRows(sqlPath: string): Promise<DashboardRow[]> {
    const sql  = fs.readFileSync(sqlPath, 'utf8');
    const pool = new Pool(this.config);
    try {
      this.logger.info('Querying EDCT data from PostgreSQL...');
      const result = await pool.query<EdctDbRow>(sql);
      this.logger.info(`  EDCT query returned ${result.rows.length} rows`);
      return result.rows.map(row => this.mapRow(row));
    } finally {
      await pool.end();
    }
  }

  private mapRow(row: EdctDbRow): DashboardRow {
    const reportingDate = dayjs(row.ReportingDate as string);
    const period = reportingDate.isValid()
      ? reportingDate.format('DD/MM/YYYY')
      : String(row.ReportingDate ?? '');

    const sex    = this.normalizeSex(String(row.Sex    ?? ''));
    const ageBand = String(row.AgeGroup ?? 'Unknown');

    return {
      Period:         period,
      State:          String(row.State    ?? ''),
      Facility:       String(row.Facility ?? ''),
      DATIMCode:      String(row.Datim    ?? ''),
      Indicator:      String(row.Section  ?? ''),
      Disaggregation: 'Sex/Age',
      Category:       sex || ageBand,
      Sex:            sex,
      AgeBand:        ageBand,
      Value:          Number(row.Value)   || 0,
      Numerator:      null,
      Denominator:    null,
      Target:         null,
      AchievementPct: null,
    };
  }

  private normalizeSex(value: string): string {
    const lower = value.toLowerCase().trim();
    if (lower === 'female') return 'Female';
    if (lower === 'male')   return 'Male';
    return value;
  }
}
