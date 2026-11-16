import { EilDateParser } from '../server/dateParser';
import { ExcelService } from '../server/excelService';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

console.log('--- TEST SUITE: EIL MANHOUR AUTOMATION & AUDIT ---');

let passed = 0;
let failed = 0;

function assert(condition: boolean, msg: string) {
  if (condition) {
    console.log(`  ✓ ${msg}`);
    passed++;
  } else {
    console.error(`  ✗ FAIL: ${msg}`);
    failed++;
  }
}

function assertEqual(actual: any, expected: any, msg: string) {
  if (actual === expected) {
    console.log(`  ✓ ${msg}`);
    passed++;
  } else {
    console.error(`  ✗ FAIL: ${msg} | Expected: ${expected}, Got: ${actual}`);
    failed++;
  }
}

// ============================================================================
// 1. DATE PARSER & REVIEW DATE EVALUATION
// ============================================================================
console.log('\n[1] Date Parser & Review Date Evaluation:');
{
  // Test Excel serial number
  const serialRes = EilDateParser.parse(45153);
  assert(serialRes.isValid, 'Excel serial number parses successfully');
  assertEqual(serialRes.formattedDate, '2023-08-15', 'Serial 45153 is 2023-08-15');

  // Test UK Date
  const ukRes = EilDateParser.parse('15/08/2026');
  assert(ukRes.isValid, '15/08/2026 parses as valid');
  assertEqual(ukRes.formattedDate, '2026-08-15', '15/08/2026 formats to 2026-08-15');

  // Test Inspection Multi-day Evaluation in Month
  const inspEval = EilDateParser.evaluateInspectionDates('2026-08-04', '2026-08-06', '2026-08', 3);
  assert(inspEval.isIncluded, 'Inspection in August 2026 is included for evaluation month 2026-08');
  assertEqual(inspEval.activeDays, 3, 'All 3 days fall within 2026-08');

  // Test Inspection Outside Month
  const inspOutside = EilDateParser.evaluateInspectionDates('2026-07-20', '2026-07-22', '2026-08', 3);
  assert(!inspOutside.isIncluded, 'Inspection in July 2026 is excluded for evaluation month 2026-08');
  assertEqual(inspOutside.activeDays, 0, '0 active days for July call evaluated in August');

  // Test QAP in Month vs Outside Month
  const qapIn = EilDateParser.evaluateQapDate('2026-08-10', '2026-08');
  assert(qapIn.isIncluded, 'QAP approved in 2026-08 is included');
  const qapOut = EilDateParser.evaluateQapDate('2025-08-10', '2026-08');
  assert(!qapOut.isIncluded, 'QAP approved in 2025 is excluded for 2026-08 evaluation');

  // Test Review Dates Evaluation
  const revSingle = EilDateParser.evaluateReviewDate('2026-10-01', '2026-10');
  assert(revSingle.isIncluded, 'Review date 2026-10-01 is included for 2026-10');
  assertEqual(revSingle.parsedDate, '2026-10-01', 'Review single date matched');

  const revComposite = EilDateParser.evaluateReviewDate('16,17-Oct-26', '2026-10');
  assert(revComposite.isIncluded, 'Review shorthand range 16,17-Oct-26 is included for 2026-10');

  const revCommaList = EilDateParser.evaluateReviewDate('12-Sep-26, 22-Oct-26', '2026-10');
  assert(revCommaList.isIncluded, 'Review composite list with October date is included for 2026-10');

  const revOutside = EilDateParser.evaluateReviewDate('15-Sep-26', '2026-10');
  assert(!revOutside.isIncluded, 'Review date in September is excluded for 2026-10 evaluation');
}

// ============================================================================
// 2. 20-POINT COMPREHENSIVE VERIFICATION FOR CLASSIFICATION & EXCEL 3 REMARKS
// ============================================================================
console.log('\n[2] Comprehensive 20-Point Verification Matrix:');
{
  const mockMapping1: any = {
    workbookType: 'excel1',
    selectedSheet: 'Sheet1',
    isValid: true,
    validationErrors: [],
    validationWarnings: [],
    fields: [
      { canonicalField: 'job_number', mappedHeader: 'Job No' },
      { canonicalField: 'balance_qty', mappedHeader: 'Balance Quantity' },
      { canonicalField: 'ocs_date', mappedHeader: 'OCS Date' },
      { canonicalField: 'head_rpo_to_job_monitor', mappedHeader: 'Head RPO' },
      { canonicalField: 'foi_number', mappedHeader: 'FOI Number' },
    ],
  };

  const mockMapping2: any = {
    workbookType: 'excel2',
    selectedSheet: 'Sheet1',
    isValid: true,
    validationErrors: [],
    validationWarnings: [],
    fields: [
      { canonicalField: 'job_number', mappedHeader: 'Job No' },
      { canonicalField: 'foi_number', mappedHeader: 'FOI Number' },
      { canonicalField: 'contractor', mappedHeader: 'Contractor/ Main Supplier' },
      { canonicalField: 'supplier', mappedHeader: 'Supplier' },
      { canonicalField: 'item', mappedHeader: 'Item' },
      { canonicalField: 'inspection_from', mappedHeader: 'Inspection Attended (From)' },
      { canonicalField: 'inspection_upto', mappedHeader: 'Inspection Attended(Upto)' },
      { canonicalField: 'review_dates', mappedHeader: 'Review dates' },
      { canonicalField: 'qap_approval', mappedHeader: 'QAP appl.' },
      { canonicalField: 'number_of_days', mappedHeader: 'No. of Days' },
    ],
  };

  const mockMapping3: any = {
    workbookType: 'excel3',
    selectedSheet: 'Sheet1',
    isValid: true,
    validationErrors: [],
    validationWarnings: [],
    fields: [
      { canonicalField: 'job_number', mappedHeader: 'Job No' },
      { canonicalField: 'foi_number', mappedHeader: 'FOI Number' },
    ],
  };

  // POINT 1: Inspection dates present -> Days × 8 goes to Inspection, 0 to Others
  const rowsP1 = [
    {
      'Job No': 'J101',
      'Contractor/ Main Supplier': 'L&T',
      'Supplier': 'Micon',
      'Item': 'Gate Valves',
      'Inspection Attended (From)': '2026-10-05',
      'Inspection Attended(Upto)': '2026-10-06',
      'Review dates': '',
      'QAP appl.': '',
      'No. of Days': 2.0,
    },
  ];
  const resP1 = ExcelService.analyzeJob('J101', '2026-10', [], mockMapping1, rowsP1, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP1.inspection, 16, 'Point 1: Inspection = 2 days × 8 = 16 MH');
  assertEqual(resP1.others, 0, 'Point 1: Others = 0 MH');

  // POINT 2: No inspection dates, valid QAP in month -> Days × 1 goes to Others
  const rowsP2 = [
    {
      'Job No': 'J102',
      'Contractor/ Main Supplier': 'Toyo',
      'Supplier': 'Sulzer',
      'Item': 'Pumps',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '',
      'QAP appl.': '2026-10-12',
      'No. of Days': 1.0,
    },
  ];
  const resP2 = ExcelService.analyzeJob('J102', '2026-10', [], mockMapping1, rowsP2, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP2.inspection, 0, 'Point 2: Inspection = 0 MH');
  assertEqual(resP2.others, 1.0, 'Point 2: Others = 1.0 MH');

  // POINT 3: No inspection dates, no QAP, valid Review dates in month -> Days × 1 goes to Others
  const rowsP3 = [
    {
      'Job No': 'J103',
      'Contractor/ Main Supplier': 'Petrofac',
      'Supplier': 'FMC',
      'Item': 'Wellhead',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-14',
      'QAP appl.': '',
      'No. of Days': 0.5,
    },
  ];
  const resP3 = ExcelService.analyzeJob('J103', '2026-10', [], mockMapping1, rowsP3, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP3.inspection, 0, 'Point 3: Inspection = 0 MH');
  assertEqual(resP3.others, 0.5, 'Point 3: Others = 0.5 MH from Review dates');

  // POINT 4: Both QAP and Review dates present without inspection -> Counted only ONCE under Others
  const rowsP4 = [
    {
      'Job No': 'J104',
      'Contractor/ Main Supplier': 'Saipem',
      'Supplier': 'Cameron',
      'Item': 'Choke Valve',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-20',
      'QAP appl.': '2026-10-20',
      'No. of Days': 1.5,
    },
  ];
  const resP4 = ExcelService.analyzeJob('J104', '2026-10', [], mockMapping1, rowsP4, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP4.inspection, 0, 'Point 4: Inspection = 0 MH');
  assertEqual(resP4.others, 1.5, 'Point 4: Others = 1.5 MH counted only once');
  assert(resP4.warnings.some((w) => w.toLowerCase().includes('qap') && w.toLowerCase().includes('review')), 'Point 4: Overlap between QAP and Review dates flagged in warnings');

  // POINT 5: Inspection dates AND QAP present -> Inspection priority, Days × 8 to Inspection, 0 to Others, conflict flagged
  const rowsP5 = [
    {
      'Job No': 'J105',
      'Contractor/ Main Supplier': 'Bechtel',
      'Supplier': 'Rotork',
      'Item': 'Actuators',
      'Inspection Attended (From)': '2026-10-10',
      'Inspection Attended(Upto)': '2026-10-11',
      'Review dates': '',
      'QAP appl.': '2026-10-10',
      'No. of Days': 2.0,
    },
  ];
  const resP5 = ExcelService.analyzeJob('J105', '2026-10', [], mockMapping1, rowsP5, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP5.inspection, 16, 'Point 5: Inspection = 16 MH (Inspection priority)');
  assertEqual(resP5.others, 0, 'Point 5: Others = 0 MH (no double count)');
  assert(resP5.warnings.some((w) => w.toLowerCase().includes('overlap') || w.toLowerCase().includes('priority')), 'Point 5: Conflict flagged when inspection & QAP both present');

  // POINT 6: Inspection dates AND Review dates present -> Inspection priority, Days × 8 to Inspection, 0 to Others, conflict flagged
  const rowsP6 = [
    {
      'Job No': 'J106',
      'Contractor/ Main Supplier': 'Fluor',
      'Supplier': 'Flowserve',
      'Item': 'Control Valves',
      'Inspection Attended (From)': '2026-10-15',
      'Inspection Attended(Upto)': '2026-10-15',
      'Review dates': '2026-10-15',
      'QAP appl.': '',
      'No. of Days': 1.0,
    },
  ];
  const resP6 = ExcelService.analyzeJob('J106', '2026-10', [], mockMapping1, rowsP6, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP6.inspection, 8, 'Point 6: Inspection = 8 MH (Inspection priority)');
  assertEqual(resP6.others, 0, 'Point 6: Others = 0 MH');
  assert(resP6.warnings.some((w) => w.toLowerCase().includes('overlap') || w.toLowerCase().includes('priority')), 'Point 6: Conflict flagged when inspection & Review dates both present');

  // POINT 7: No inspection, no QAP, no Review dates -> 0 contribution (NO_ACTIVITY)
  const rowsP7 = [
    {
      'Job No': 'J107',
      'Contractor/ Main Supplier': 'Technip',
      'Supplier': 'BHEL',
      'Item': 'Turbine',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '',
      'QAP appl.': '',
      'No. of Days': 1.0,
    },
  ];
  const resP7 = ExcelService.analyzeJob('J107', '2026-10', [], mockMapping1, rowsP7, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP7.inspection, 0, 'Point 7: Inspection = 0 MH');
  assertEqual(resP7.others, 0, 'Point 7: Others = 0 MH');
  assertEqual(resP7.sourceCallsExcel2[0]?.status, 'NO_ACTIVITY', 'Point 7: Status is NO_ACTIVITY');

  // POINT 8: Decimal Number of Days (0.5, 1.5, 0.25) preserved exactly and never multiplied by 8 for Others
  const rowsP8 = [
    {
      'Job No': 'J108',
      'Contractor/ Main Supplier': 'L&T',
      'Supplier': 'Micon',
      'Item': 'Valves',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-22',
      'QAP appl.': '',
      'No. of Days': 0.25,
    },
  ];
  const resP8 = ExcelService.analyzeJob('J108', '2026-10', [], mockMapping1, rowsP8, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP8.others, 0.25, 'Point 8: Decimal 0.25 preserved exactly in Others');

  // POINT 9: Multiple eligible records for one job -> aggregated correctly
  const rowsP9 = [
    {
      'Job No': 'J109',
      'Contractor/ Main Supplier': 'Main Contractor A',
      'Supplier': 'Vendor X',
      'Item': 'Motors',
      'Inspection Attended (From)': '2026-10-07',
      'Inspection Attended(Upto)': '2026-10-10',
      'Review dates': '',
      'QAP appl.': '',
      'No. of Days': 4.0, // 4 * 8 = 32 MH
    },
    {
      'Job No': 'J109',
      'Contractor/ Main Supplier': 'Main Contractor A',
      'Supplier': 'Vendor Y',
      'Item': 'Cables',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '',
      'QAP appl.': '2026-10-18',
      'No. of Days': 1.0, // 1 MH
    },
    {
      'Job No': 'J109',
      'Contractor/ Main Supplier': '',
      'Supplier': 'Vendor Z',
      'Item': 'Transformers',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-25',
      'QAP appl.': '',
      'No. of Days': 0.5, // 0.5 MH
    },
  ];
  const resP9 = ExcelService.analyzeJob('J109', '2026-10', [], mockMapping1, rowsP9, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP9.inspection, 32, 'Point 9: Aggregated Inspection = 32 MH');
  assertEqual(resP9.others, 1.5, 'Point 9: Aggregated Others = 1.0 + 0.5 = 1.5 MH');
  assertEqual(resP9.total, 33.5, 'Point 9: Aggregated Total = 33.5 MH');

  // POINT 13: Historical month records excluded from calculations and remarks
  const rowsP13 = [
    {
      'Job No': 'J113',
      'Contractor/ Main Supplier': 'Historical Ctr',
      'Supplier': 'Historical Sup',
      'Item': 'Historical Item',
      'Inspection Attended (From)': '2025-05-10',
      'Inspection Attended(Upto)': '2025-05-12',
      'Review dates': '',
      'QAP appl.': '',
      'No. of Days': 3.0,
    },
  ];
  const resP13 = ExcelService.analyzeJob('J113', '2026-10', [], mockMapping1, rowsP13, mockMapping2, [], mockMapping3, 0);
  assertEqual(resP13.inspection, 0, 'Point 13: Historical inspection excluded from 2026-10');
  assertEqual(resP13.others, 0, 'Point 13: Historical others excluded from 2026-10');
  assertEqual(resP13.sourceCallsExcel2[0]?.eligibility, 'Excluded', 'Point 13: Historical record marked Excluded');

  // POINT 14 & 15: Authoritative Job Discovery (J1 UNION J2)
  const excel1Jobs = [{ 'Job No': 'B114', 'FOI Number': 'FOI-114', 'Balance Quantity': 2 }];
  const excel2Jobs = [
    {
      'Job No': 'B115',
      'FOI Number': 'FOI-115',
      'Inspection Attended (From)': '2026-10-02',
      'Inspection Attended(Upto)': '2026-10-02',
      'No. of Days': 1.0,
    },
    {
      'Job No': 'B999',
      'FOI Number': 'FOI-999',
      'Inspection Attended (From)': '2025-01-01',
      'Inspection Attended(Upto)': '2025-01-02',
      'No. of Days': 2.0,
    },
  ];

  const wb1Recon = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb1Recon, XLSX.utils.json_to_sheet(excel1Jobs), 'Sheet1');
  const buf1Recon = XLSX.write(wb1Recon, { type: 'buffer', bookType: 'xlsx' });

  const wb2Recon = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2Recon, XLSX.utils.json_to_sheet(excel2Jobs), 'Sheet1');
  const buf2Recon = XLSX.write(wb2Recon, { type: 'buffer', bookType: 'xlsx' });

  const wb3Recon = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb3Recon, XLSX.utils.json_to_sheet([{ 'Job No': 'B114', 'FOI Number': 'FOI-114' }]), 'Sheet1');
  const buf3Recon = XLSX.write(wb3Recon, { type: 'buffer', bookType: 'xlsx' });

  const reconTest = ExcelService.reconcileJobs(
    buf1Recon,
    mockMapping1,
    buf2Recon,
    mockMapping2,
    buf3Recon,
    mockMapping3,
    '2026-10'
  );
  const discoveredJobNos = reconTest.jobs.map((j) => j.jobNumber);
  assert(discoveredJobNos.includes('B114'), 'Point 14: Excel 1 job included');
  assert(discoveredJobNos.includes('B115'), 'Point 14: PM Call Logs current-month job included (Case C)');
  assert(!discoveredJobNos.includes('B999'), 'Point 15: PM Call Logs historical-only job excluded (Case D)');

  // POINT 16: Existing Excel 1 rules (Orders for FD strict logic & Expediting formula)
  const rowsExcel1 = [
    { 'Job No': 'J116', 'Balance Quantity': 3, 'OCS Date': '', 'Head RPO': '' }, // Running Order (Balance > 0)
    { 'Job No': 'J116', 'Balance Quantity': 0, 'OCS Date': '2026-10-12', 'Head RPO': '' }, // OCS Done (in month)
    { 'Job No': 'J116', 'Balance Quantity': 0, 'OCS Date': '', 'Head RPO': '' }, // Orders for FD (Balance = 0, empty OCS Date, blank Head RPO)
  ];
  const resP16 = ExcelService.analyzeJob('J116', '2026-10', rowsExcel1, mockMapping1, [], mockMapping2, [], mockMapping3, 0);
  assertEqual(resP16.runningOrders, 1, 'Point 16: Running Orders = 1');
  assertEqual(resP16.ocsDone, 1, 'Point 16: OCS Done = 1');
  assertEqual(resP16.fd, 1, 'Point 16: Orders for FD = 1 (Balance 0, OCS Date empty, Head RPO blank)');
  assertEqual(resP16.expediting, 4, 'Point 16: Expediting = 2 × (1 + 1) = 4 MH');

  // POINT 21: Scheduled Inspection Date (From)/(Upto) present while Attended Inspection & QAP absent -> Review Date STILL counts in Others
  const rowsScheduledOnly = [
    {
      'Job No': 'B957',
      'Contractor/ Main Supplier': 'Megha(Pkg09)',
      'Supplier': 'Micon',
      'Item': 'Valves',
      'Inspection Date (From)': '2026-10-01',
      'Inspection Date(Upto)': '2026-10-01',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-01',
      'QAP appl.': '-',
      'No. of Days': 0.5,
    },
    {
      'Job No': 'B957',
      'Contractor/ Main Supplier': 'Megha(Pkg09)',
      'Supplier': 'Micon',
      'Item': 'Valves',
      'Inspection Date (From)': '2026-09-10',
      'Inspection Date(Upto)': '2026-09-10',
      'Inspection Attended (From)': '',
      'Inspection Attended(Upto)': '',
      'Review dates': '2026-10-22',
      'QAP appl.': 'N/A',
      'No. of Days': 0.5,
    },
  ];
  const resSched = ExcelService.analyzeJob('B957', '2026-10', [], mockMapping1, rowsScheduledOnly, mockMapping2, [], mockMapping3, 0);
  assertEqual(resSched.inspection, 0, 'Point 21: Scheduled call dates without Attended dates give 0 Inspection MH');
  assertEqual(resSched.others, 1.0, 'Point 21: Review dates counted in Others (0.5 + 0.5 = 1.0 MH) despite scheduled dates and "-"/"N/A" QAP');
  const elig0 = ExcelService.evaluateExcel2RowEligibility(rowsScheduledOnly[0], mockMapping2, '2026-10');
  assertEqual(elig0.type, 'REVIEW', 'Point 21: evaluateExcel2RowEligibility classifies scheduled-only + Review row as REVIEW');
  assertEqual(elig0.computedMh, 0.5, 'Point 21: evaluateExcel2RowEligibility computes 0.5 MH');

  // POINT 22: Month-boundary spanning inspections preserve full authoritative No. of Days × 8 MH (no pro-rating)
  const sepOctSpan = EilDateParser.evaluateInspectionDates('28 Sep 26', '1 Oct 26', '2026-10', 4);
  assert(sepOctSpan.isIncluded, 'Point 22: 28 Sep 26 -> 1 Oct 26 included in 2026-10');
  assertEqual(sepOctSpan.activeDays, 4, 'Point 22: Full 4 days preserved for 28 Sep 26 -> 1 Oct 26 (no pro-rating)');
  assertEqual(sepOctSpan.status, 'FULL', 'Point 22: Status is FULL for month-boundary inspection');

  const octNovSpan = EilDateParser.evaluateInspectionDates('30 Oct 26', '2 Nov 26', '2026-10', 4);
  assert(octNovSpan.isIncluded, 'Point 22: 30 Oct 26 -> 2 Nov 26 included in 2026-10');
  assertEqual(octNovSpan.activeDays, 4, 'Point 22: Full 4 days preserved for 30 Oct 26 -> 2 Nov 26');

  const rowsB895Regression = [
    {
      'Job No': 'B895',
      'S.No': 1824,
      'Contractor/ Main Supplier': 'Toyo',
      'Supplier': 'Schneider',
      'Item': 'Sw Bds',
      'Inspection Attended (From)': '28 Sep 26',
      'Inspection Attended(Upto)': '1 Oct 26',
      'Review dates': '-',
      'QAP appl.': '-',
      'No. of Days': 4,
    },
    {
      'Job No': 'B895',
      'S.No': 1851,
      'Contractor/ Main Supplier': 'Toyo',
      'Supplier': 'Belgaum Aqua',
      'Item': 'Valves',
      'Inspection Attended (From)': '15 Oct 26',
      'Inspection Attended(Upto)': '-',
      'Review dates': '-',
      'QAP appl.': '-',
      'No. of Days': 1,
    },
  ];
  const resB895 = ExcelService.analyzeJob('B895', '2026-10', [], mockMapping1, rowsB895Regression, mockMapping2, [], mockMapping3, 0);
  assertEqual(resB895.inspection, 40, 'Point 22: B895 Call #1824 (4d × 8 = 32) + Call #1851 (1d × 8 = 8) = 40 MH');
  assertEqual(resB895.sourceCallsExcel2[0]?.computedMh, 32, 'Point 22: Call #1824 computedMh is 32 MH (4 × 8)');
  assertEqual(resB895.sourceCallsExcel2[1]?.computedMh, 8, 'Point 22: Call #1851 computedMh is 8 MH (1 × 8)');

  const b795Elig = ExcelService.evaluateExcel2RowEligibility(
    {
      'Job No': 'B795',
      'Inspection Attended (From)': '30-Sep-26',
      'Inspection Attended(Upto)': '1-Oct-26',
      'No. of Days': 2,
    },
    mockMapping2,
    '2026-10'
  );
  assertEqual(b795Elig.type, 'INSPECTION', 'Point 22: B795 30-Sep-26 -> 1-Oct-26 classified as INSPECTION');
  assertEqual(b795Elig.days, 2, 'Point 22: B795 full 2 days preserved');
  assertEqual(b795Elig.computedMh, 16, 'Point 22: B795 computedMh is 16 MH (2 × 8)');
}

// ============================================================================
// 3. EXCEL 3 OUTPUT GENERATION, REMARKS (POINTS 10, 11, 12, 17, 18, 19, 20)
// ============================================================================
console.log('\n[3] Excel 3 Output Workbook Generation & Formatted Remarks Verification:');
async function testExcel3Generation() {
  const samples = ExcelService.createSampleWorkbooks();
  const i3 = ExcelService.parseWorkbookBuffer(samples.excel3.buffer, samples.excel3.fileName, 'excel3');
  const m3 = ExcelService.autoMapHeaders('excel3', i3.sheets[0]);

  // Create an approved review object for Job B269 with multiple calls including missing contractor
  const approvedReviews: Record<string, any> = {
    'B269': {
      jobNumber: 'B269',
      foiNumber: 'FOI-2691',
      origin: 'Both',
      runningOrders: 1,
      fd: 0,
      ocsDone: 1,
      expediting: 4,
      inspection: 16,
      others: 1.5,
      meeting: 0,
      total: 21.5,
      approvalStatus: 'APPROVED',
      isOverridden: false,
      sourceCallsExcel2: [
        {
          callNo: 1,
          contractor: 'Larsen & Toubro',
          vendor: 'Kirloskar Brothers Ltd',
          itemDescription: 'Centrifugal Process Pump',
          numberOfDays: 2.0,
          eligibility: 'Eligible',
          status: 'INCLUDED_INSPECTION',
        },
        {
          callNo: 2,
          contractor: 'Larsen & Toubro',
          vendor: 'Sulzer Pumps India',
          itemDescription: 'Mechanical Seals & Spares',
          numberOfDays: 0.5,
          eligibility: 'Eligible',
          status: 'INCLUDED_QAP',
        },
        {
          callNo: 21,
          contractor: 'larsen & toubro',
          vendor: 'Sulzer Pumps India',
          itemDescription: 'Mechanical Seals & Spares',
          numberOfDays: 0.5,
          eligibility: 'Eligible',
          status: 'INCLUDED_REVIEW',
        },
        {
          callNo: 3,
          contractor: '', // Intentionally blank to test 'Not Available'
          vendor: 'EagleBurgmann',
          itemDescription: 'Gaskets',
          numberOfDays: 0.5,
          eligibility: 'Eligible',
          status: 'INCLUDED_REVIEW',
        },
        {
          callNo: 4,
          contractor: 'Old Contractor',
          vendor: 'Old Vendor',
          itemDescription: 'Old Item',
          numberOfDays: 2.0,
          eligibility: 'Excluded', // Excluded call from historical period
          status: 'EXCLUDED_DATES',
        },
      ],
    },
    'G660': {
      jobNumber: 'G660',
      foiNumber: 'FOI-6601',
      origin: 'PM Call Logs',
      runningOrders: 0,
      fd: 0,
      ocsDone: 0,
      expediting: 0,
      inspection: 16,
      others: 0,
      meeting: 2,
      total: 18,
      approvalStatus: 'APPROVED',
      isOverridden: false,
      sourceCallsExcel2: [
        {
          callNo: 10,
          contractor: 'Samsung Engineering',
          vendor: 'Technip Energies',
          itemDescription: 'Furnace Burners',
          numberOfDays: 2.0,
          eligibility: 'Eligible',
          status: 'INCLUDED_INSPECTION',
        },
      ],
    },
  };

  const genResult = await ExcelService.generateOutputWorkbook(
    samples.excel3.buffer,
    m3,
    approvedReviews,
    '2026-08'
  );

  assert(genResult.buffer.length > 0, 'Point 20: Output buffer generated successfully');
  assertEqual(genResult.report.status, 'PASS', 'Point 20: Output report status is PASS');

  // Physically inspect the generated ExcelJS workbook
  const verifyWb = new ExcelJS.Workbook();
  await verifyWb.xlsx.load(genResult.buffer);

  const mainSheet = verifyWb.worksheets[0];
  assert(mainSheet !== undefined, 'Point 20: Main output worksheet exists');

  // Verify headers: 12 columns
  const headerRow = mainSheet.getRow(3);
  assertEqual(headerRow.getCell(1).value, 'S.No', 'Point 18: Col 1 is S.No');
  assertEqual(headerRow.getCell(2).value, 'Job No', 'Point 18: Col 2 is Job No');
  assertEqual(headerRow.getCell(3).value, 'FOI Number', 'Point 18: Col 3 is FOI Number');
  assertEqual(headerRow.getCell(4).value, 'No. of Running orders', 'Col 4 is No. of Running orders');
  assertEqual(headerRow.getCell(5).value, 'Orders for FD f/up', 'Col 5 is Orders for FD f/up');
  assertEqual(headerRow.getCell(6).value, 'OCS done', 'Col 6 is OCS done');
  assertEqual(headerRow.getCell(7).value, 'Exp.', 'Col 7 is Exp.');
  assertEqual(headerRow.getCell(8).value, 'Inspn', 'Col 8 is Inspn');
  assertEqual(headerRow.getCell(9).value, 'Others', 'Col 9 is Others');
  assertEqual(headerRow.getCell(10).value, 'Meeting', 'Col 10 is Meeting');
  assertEqual(headerRow.getCell(11).value, 'Total', 'Col 11 is Total');
  assertEqual(headerRow.getCell(12).value, '*Others/ Remarks', 'Point 10: Col 12 is *Others/ Remarks');

  // Check row 4 (B269)
  const r4 = mainSheet.getRow(4);
  assertEqual(r4.getCell(1).value, 1, 'Row 4 S.No is 1');
  assertEqual(r4.getCell(2).value, 'B269', 'Row 4 Job No is B269');
  assertEqual(r4.getCell(3).value, 'FOI-2691', 'Point 18: Row 4 FOI Number is FOI-2691');
  assertEqual(r4.getCell(7).value, 4, 'Row 4 Expediting is 4');
  assertEqual(r4.getCell(8).value, 16, 'Row 4 Inspection is 16');
  assertEqual(r4.getCell(9).value, 1.5, 'Row 4 Others is 1.5');
  const r4Formula = (r4.getCell(11).value as any)?.formula;
  assertEqual(r4Formula, 'G4+H4+I4+J4', 'Point 17: Row 4 Total has live formula =G4+H4+I4+J4');

  // Check Column 12 (*Others/ Remarks) for B269
  const r4Remarks = String(r4.getCell(12).value ?? '');
  const r4RemarksLines = r4Remarks.split('\n');

  // POINT 10: Only Others (QAP/Review) records appear, merged by (Contractor, Supplier, Item) with summed Days
  assertEqual(r4RemarksLines.length, 2, 'Point 10: Exactly 2 merged Others lines appear in Column 12 (Inspection & Excluded omitted, duplicates merged)');

  // POINT 11: Exact format Contractor: ... | Supplier: ... | Item: ... | Days: ... (0.5 + 0.5 = 1)
  assertEqual(
    r4RemarksLines[0],
    'Contractor: Larsen & Toubro | Supplier: Sulzer Pumps India | Item: Mechanical Seals & Spares | Days: 1',
    'Point 11: Merged QAP+Review entry sums Days (0.5 + 0.5 = 1) without duplicating Contractor/Supplier/Item'
  );

  // POINT 12: Missing Contractor displays 'Not Available'
  assertEqual(
    r4RemarksLines[1],
    'Contractor: Not Available | Supplier: EagleBurgmann | Item: Gaskets | Days: 0.5',
    'Point 12: Call 3 with blank contractor displays "Contractor: Not Available"'
  );

  // Check text wrap and dynamic row height on Row 4
  assert(r4.getCell(12).alignment?.wrapText === true, 'Point 20: Cell 12 has wrapText enabled');
  assert((r4.height ?? 0) >= 35, `Point 20: Row 4 height is dynamically adjusted (height: ${r4.height})`);

  // Check Column 12 width
  const col12 = mainSheet.getColumn(12);
  assert((col12.width ?? 0) >= 50, `Point 20: Column 12 width is expanded (width: ${col12.width})`);

  // Check row 5 (G660 - Inspection only, no Others calls -> clean blank Column 12)
  const r5 = mainSheet.getRow(5);
  assertEqual(r5.getCell(1).value, 2, 'Row 5 S.No is 2');
  assertEqual(r5.getCell(2).value, 'G660', 'Row 5 Job No is G660');
  assertEqual(r5.getCell(3).value, 'FOI-6601', 'Point 18: Row 5 FOI Number is FOI-6601');
  assertEqual(
    String(r5.getCell(12).value ?? ''),
    '',
    'Point 11: Row 5 G660 (Inspection-only) omits inspection lines and boilerplate from *Others/ Remarks'
  );

  // Check row 6 (Leave row)
  const r6 = mainSheet.getRow(6);
  assertEqual(r6.getCell(2).value, 'Leave', 'Point 19: Row 6 is Leave row');
  assertEqual(r6.getCell(10).value, 8, 'Point 19: Row 6 Meeting (Leave) is 8');
  assertEqual(r6.getCell(12).value, 'Annual Leave', 'Point 19: Row 6 Remarks is Annual Leave');

  // Check row 7 (TOTAL row)
  const r7 = mainSheet.getRow(7);
  assertEqual(r7.getCell(2).value, 'TOTAL', 'Point 19: Row 7 is TOTAL row');
  const sumExpFormula = (r7.getCell(7).value as any)?.formula;
  assertEqual(sumExpFormula, 'SUM(G4:G6)', 'Point 19: TOTAL row has live formula SUM(G4:G6) summing jobs + Leave');

  // Check EIL_Audit_Trail sheet
  const auditSheet = verifyWb.getWorksheet('EIL_Audit_Trail');
  assert(auditSheet !== undefined, 'Point 20: EIL_Audit_Trail sheet exists and is appended');
}

testExcel3Generation().then(() => {
  console.log(`\n======================================================`);
  console.log(`Test Execution Finished: ${passed} passed, ${failed} failed.`);
  console.log(`======================================================\n`);
  if (failed > 0) {
    process.exit(1);
  } else {
    process.exit(0);
  }
}).catch((err) => {
  console.error('Unhandled failure:', err);
  process.exit(1);
});
