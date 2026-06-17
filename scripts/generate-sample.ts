/**
 * generate-sample.ts
 * Generates a sample RADET.xlsx with realistic test data for engine testing.
 * Run with: npx ts-node scripts/generate-sample.ts
 */
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';
import dayjs from 'dayjs';

const TODAY = dayjs().format('YYYY-MM-DD');
const sheetHeadersPath = path.resolve(__dirname, '../config/sheetHeaders.yaml');
const sheetHeaders = yaml.load(fs.readFileSync(sheetHeadersPath, 'utf8')) as Record<string, string[]>;
const COMBINED_RADET_HEADERS = sheetHeaders['CombinedRADET'] ?? [];
const COMBINE_HTS_HEADERS = sheetHeaders['CombineHTS'] ?? [];

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
const ART_STATUSES = [
  'Active',
  'Lost to Follow-Up',
  'Transferred Out',
  'Stopped Treatment',
  'Dead',
  'Inactive',
];
const ART_SETTINGS = ['51', '52', 'Facility', 'COMMUNITY'];

function randomARTStatus(): string {
  const r = Math.random();
  if (r < 0.50) return 'Active';
  if (r < 0.754) return 'Lost to Follow-Up';
  if (r < 0.856) return 'Transferred Out';
  if (r < 0.932) return 'Stopped Treatment';
  if (r < 0.962) return 'Dead';
  return 'Inactive';
}

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
    const isToday = Math.random() > 0.3; // 70% chance record is today
    
    // Generate Last Pickup Date and Months of ARV Refill
    // For some records, calculate so expected refill date = today
    let lastPickupDate: string;
    let monthsOfARVRefill: number;
    
    if (isToday && Math.random() > 0.5) {
      // 50% of "today" records: expected refill date IS today
      monthsOfARVRefill = Math.floor(Math.random() * 6) + 1; // 1-6 months
      const pickupDay = dayjs().subtract(monthsOfARVRefill, 'months');
      lastPickupDate = pickupDay.format('YYYY-MM-DD');
    } else {
      // Random past pickup date with random refill months
      lastPickupDate = randomDate(180);
      monthsOfARVRefill = Math.floor(Math.random() * 6) + 1;
    }
    
    // Create object with keys in exact order per sheetHeaders.yaml
    const row: Record<string, unknown> = {};
    row.State = fac.state;
    row.LGA = fac.lga;
    row.LGAOfResidence = fac.lga;
    row.Facility = fac.name;
    row.DATIMCode = fac.datim;
    row.PatientId = `RAD-${String(i + 1).padStart(5, '0')}`;
    row.NDRPatientId = '';
    row.HospitalNumber = '';
    row.UniqueId = '';
    row.HouseholdUniqueNo = '';
    row.OVCUniqueId = '';
    row.Sex = randomFrom(SEXES);
    row.TargetGroup = '';
    row.CurrentWeight = '';
    row.PregnancyStatus = '';
    row.DateOfBirth = '';
    row.Age = randomFrom(AGES);
    row.CareEntryPoint = randomFrom(CARE_ENTRY_POINTS);
    row.DateOfRegistration = '';
    row.EnrollmentDate = isToday ? TODAY : randomDate(90);
    row.ARTStartDate = isToday ? TODAY : randomDate(90);
    row.LastPickupDate = lastPickupDate;
    row.MonthsOfARVRefill = monthsOfARVRefill;
    row.RegimenLineAtARTStart = '';
    row.RegimenAtARTStart = '';
    row.DateOfStartOfCurrentARTRegimen = '';
    row.CurrentRegimenLine = '';
    row.CurrentARTRegimen = '';
    row.ClinicalStagingAtLastVisit = '';
    row.DateOfLastCD4Count = '';
    row.LastCD4Count = '';
    row.DateOfVLSampleCollection = '';
    row.DateOfCurrentVLResultSample = '';
    // Generate viral load data: ~40% have VL results spread across fiscal year
    // Mix of formats: plain numbers, <50 format, and NotDetected
    if (Math.random() > 0.6) {
      // VL dates spread across 365 days (entire year)
      const vlDaysBack = Math.floor(Math.random() * 365);
      row.DateOfCurrentViralLoad = dayjs().subtract(vlDaysBack, 'days').format('YYYY-MM-DD');
      const vlRandom = Math.random();
      if (vlRandom < 0.15) {
        // 15% NotDetected (undetectable)
        row.CurrentViralLoad = 'NotDetected';
      } else if (vlRandom < 0.30) {
        // 15% <50 (very low, undetectable range)
        row.CurrentViralLoad = '<50';
      } else if (vlRandom < 0.55) {
        // 25% suppressed numbers (50-1000)
        row.CurrentViralLoad = String(Math.floor(Math.random() * 950) + 50);
      } else {
        // 45% unsuppressed (>1000)
        row.CurrentViralLoad = String(Math.floor(Math.random() * 500000) + 1001);
      }
    } else {
      row.CurrentViralLoad = '';
      row.DateOfCurrentViralLoad = '';
    }
    row.DateOfCommencementOfEAC = '';
    row.ViralLoadIndication = '';
    row.ViralLoadEligibilityStatus = '';
    row.DateOfVLEligibilityStatus = '';
    row.CurrentARTStatus = randomARTStatus();
    if (row.CurrentARTStatus === 'Active') {
      const currentVlNum = parseFloat(String(row.CurrentViralLoad));
      if (!Number.isNaN(currentVlNum) && currentVlNum > 1000 && row.DateOfCurrentViralLoad) {
        const startDate = dayjs(String(row.DateOfCurrentViralLoad)).add(7, 'day');
        row.DateOfCommencementOfEAC = startDate.isAfter(dayjs())
          ? String(row.DateOfCurrentViralLoad)
          : startDate.format('YYYY-MM-DD');
      }
    }
    row.DateOfCurrentARTStatus = '';
    row.ClientVerificationOutcome = '';
    row.CauseOfDeath = '';
    row.VACauseOfDeath = '';
    row.PreviousARTStatus = '';
    row.ConfirmedDateOfPreviousARTStatus = '';
    row.DateOfTBScreening = isToday ? TODAY : randomDate(30);
    row.TBStatus = randomFrom(TB_STATUSES);
    row.DateOfTBSampleCollection = isToday ? TODAY : randomDate(14);
    row.DateOfTBResultReturn = isToday ? TODAY : randomDate(7);
    row.TBTreatmentStartDate = isToday ? TODAY : randomDate(7);
    row.DateOfTBTreatmentOutcome = '';
    row.TBTreatmentOutcome = '';
    row.ARTEnrollmentSetting = randomFrom(ART_SETTINGS);
    rows.push(row);
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
    for (const row of radetRows) {
      radetWs.addRow(COMBINED_RADET_HEADERS.map((h) => row[h] ?? ''));
    }
  }

  // ── CombineHTS sheet ──────────────────────────────────────────────────
  const htsWs = wb.addWorksheet('CombineHTS');
  const htsRows = generateHTSRows(300);
  if (htsRows.length > 0) {
    for (const row of htsRows) {
      htsWs.addRow(COMBINE_HTS_HEADERS.map((h) => row[h] ?? ''));
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
