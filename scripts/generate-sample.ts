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
    row.CurrentViralLoad = '';
    row.DateOfCurrentViralLoad = '';
    row.ViralLoadIndication = '';
    row.ViralLoadEligibilityStatus = '';
    row.DateOfVLEligibilityStatus = '';
    row.CurrentARTStatus = randomFrom(ART_STATUSES);
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
