/**
 * generate-sample.ts
 * Generates a sample RADET.xlsx with realistic test data for engine testing.
 * Run with: npx ts-node scripts/generate-sample.ts
 */
import ExcelJS from 'exceljs';
import path from 'path';
import dayjs from 'dayjs';

const TODAY = dayjs().format('YYYY-MM-DD');

const FACILITIES = [
  { name: 'General Hospital Wukari', datim: 'JPBcTpp6XUu', lga: 'Wukari', state: 'Taraba' },
  { name: 'PHC Ibi', datim: 'KQmcBpp7YUv', lga: 'Ibi', state: 'Taraba' },
  { name: 'Fed Medical Centre Jalingo', datim: 'LRndCpp8ZVw', lga: 'Jalingo', state: 'Taraba' },
  { name: 'Cottage Hospital Yorro', datim: 'MSoeDqq9AWx', lga: 'Yorro', state: 'Taraba' },
];

const SEXES = ['Male', 'Female'];
const AGES = [2, 7, 12, 17, 22, 27, 32, 37, 42, 47, 52, 58];
const CARE_ENTRY_POINTS = ['OPD', 'ANC', 'TB Clinic', 'Emergency', 'PMTCT', 'Self Referral'];
const TB_STATUSES = [
  'Presumptive TB', 'Presumptive', 'TB/HIV co-infected', 'TB suspect',
  'Not a TB Case', 'TB Treatment Completed',
];
const ART_STATUSES = ['Active', 'Inactive', 'LTFU', 'Dead'];
const ART_SETTINGS = ['51', '52', 'Facility', 'COMMUNITY'];

function randomFrom<T>(arr: T[]): T {
  return arr[Math.floor(Math.random() * arr.length)];
}

function randomDate(daysBack: number): string {
  const d = dayjs().subtract(Math.floor(Math.random() * daysBack), 'day');
  return d.format('YYYY-MM-DD');
}

function generateRADETRows(count: number): Record<string, unknown>[] {
  const rows: Record<string, unknown>[] = [];
  for (let i = 0; i < count; i++) {
    const fac = randomFrom(FACILITIES);
    const isToday = Math.random() > 0.7; // 30% chance record is today
    rows.push({
      PatientID: `RAD-${String(i + 1).padStart(5, '0')}`,
      Facility: fac.name,
      DATIMCode: fac.datim,
      LGA: fac.lga,
      State: fac.state,
      Sex: randomFrom(SEXES),
      Age: randomFrom(AGES),
      EnrollmentDate: isToday ? TODAY : randomDate(90),
      ARTStartDate: isToday ? TODAY : randomDate(90),
      CareEntryPoint: randomFrom(CARE_ENTRY_POINTS),
      CurrentARTStatus: randomFrom(ART_STATUSES),
      TBStatus: randomFrom(TB_STATUSES),
      DateOfTBScreening: isToday ? TODAY : randomDate(30),
      DateOfTBSampleCollection: isToday ? TODAY : randomDate(14),
      DateOfTBResultReturn: isToday ? TODAY : randomDate(7),
      TBTreatmentStartDate: isToday ? TODAY : randomDate(7),
      ARTEnrollmentSetting: randomFrom(ART_SETTINGS),
    });
  }
  return rows;
}

function generateHTSRows(count: number): Record<string, unknown>[] {
  const rows: Record<string, unknown>[] = [];
  for (let i = 0; i < count; i++) {
    const fac = randomFrom(FACILITIES);
    const isPos = Math.random() > 0.85; // ~15% positivity
    const isToday = Math.random() > 0.6;
    rows.push({
      ClientID: `HTS-${String(i + 1).padStart(5, '0')}`,
      Facility: fac.name,
      DATIMCode: fac.datim,
      LGA: fac.lga,
      State: fac.state,
      Sex: randomFrom(SEXES),
      Age: randomFrom(AGES),
      finalHIVTestResult: isPos ? 'Positive' : 'Negative',
      dateOfHIVTesting: isToday ? TODAY : randomDate(30),
    });
  }
  return rows;
}

async function main(): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // ── CombinedRADET sheet ────────────────────────────────────────────────
  const radetWs = wb.addWorksheet('CombinedRADET');
  const radetRows = generateRADETRows(500);
  if (radetRows.length > 0) {
    const headers = Object.keys(radetRows[0]);
    radetWs.addRow(headers);
    for (const row of radetRows) {
      radetWs.addRow(headers.map((h) => row[h]));
    }
  }

  // ── CombineHTS sheet ──────────────────────────────────────────────────
  const htsWs = wb.addWorksheet('CombineHTS');
  const htsRows = generateHTSRows(300);
  if (htsRows.length > 0) {
    const headers = Object.keys(htsRows[0]);
    htsWs.addRow(headers);
    for (const row of htsRows) {
      htsWs.addRow(headers.map((h) => row[h]));
    }
  }

  const outPath = path.resolve(__dirname, '../input/RADET.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`\n✅ Sample RADET.xlsx generated: ${outPath}`);
  console.log(`   CombinedRADET: ${radetRows.length} rows`);
  console.log(`   CombineHTS:    ${htsRows.length} rows`);
  console.log('\nNow run: npm run dev\n');
}

main().catch((e) => { console.error(e); process.exit(1); });
