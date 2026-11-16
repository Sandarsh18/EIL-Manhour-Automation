import * as XLSX from 'xlsx';

async function runTests() {
  console.log('====================================================');
  console.log('STARTING STRICT VERIFICATION OF DEFECTS & API ROUTES');
  console.log('====================================================\n');

  const BASE_URL = 'http://localhost:3000';
  let passedCount = 0;
  let failedCount = 0;

  function assert(condition: boolean, testName: string, details?: string) {
    if (condition) {
      console.log(`[PASS] ${testName}`);
      passedCount++;
    } else {
      console.error(`[FAIL] ${testName} - ${details || 'Assertion failed'}`);
      failedCount++;
    }
  }

  // Helper to test an endpoint for JSON response
  async function checkJsonResponse(res: Response, endpointDesc: string) {
    const contentType = res.headers.get('content-type') || '';
    const isJson = contentType.includes('application/json');
    assert(isJson, `${endpointDesc} returns Content-Type application/json (got: ${contentType})`);
    const text = await res.text();
    const hasHtmlTag = text.toLowerCase().includes('<!doctype') || text.toLowerCase().includes('<html');
    assert(!hasHtmlTag, `${endpointDesc} body does NOT contain HTML (no '<!doctype' or '<html')`);
    
    let parsedJson: any = null;
    try {
      parsedJson = JSON.parse(text);
      assert(true, `${endpointDesc} is valid JSON with no "Unexpected token <" error`);
    } catch (e: any) {
      assert(false, `${endpointDesc} failed JSON.parse: ${e.message}`, text.slice(0, 100));
    }
    return parsedJson;
  }

  // ----------------------------------------------------
  // TEST SECTION 1: Step 1 & Step 2 API Call Verification
  // ----------------------------------------------------
  console.log('--- SECTION 1: Testing All Step 1 & Step 2 API Routes ---');

  // 1.1 GET /api/workbooks
  let res = await fetch(`${BASE_URL}/api/workbooks`);
  let json = await checkJsonResponse(res, 'GET /api/workbooks');
  assert(json.success === true, 'GET /api/workbooks returned success: true');

  // 1.2 POST /api/workbooks/load-samples
  res = await fetch(`${BASE_URL}/api/workbooks/load-samples`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/workbooks/load-samples');
  assert(json.success === true && json.data.excel1 && json.data.excel2 && json.data.excel3, 'Samples loaded with all 3 workbooks');

  // Run analysis so we have downstream calculations to test invalidation later
  await fetch(`${BASE_URL}/api/analysis/run`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ evaluationMonth: '2026-08' }),
  });

  // Verify calculations and reviews exist
  res = await fetch(`${BASE_URL}/api/analysis/results`);
  json = await checkJsonResponse(res, 'GET /api/analysis/results');
  const calcKeysBefore = Object.keys(json.data || {});
  assert(calcKeysBefore.length > 0, `Active calculations exist before replacement test (${calcKeysBefore.length} jobs)`);

  // 1.3 GET /api/mappings
  res = await fetch(`${BASE_URL}/api/mappings`);
  json = await checkJsonResponse(res, 'GET /api/mappings');
  assert(json.success === true && json.data.excel1 && json.data.excel2 && json.data.excel3, 'GET /api/mappings returned mappings for all workbooks');

  // 1.4 POST /api/mappings/validate
  res = await fetch(`${BASE_URL}/api/mappings/validate`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/mappings/validate');
  assert(json.success === true && json.data.isValid === true, 'POST /api/mappings/validate validates authentic mappings as valid');

  // 1.5 POST /api/mappings/excel1 (Update dynamic mapping)
  res = await fetch(`${BASE_URL}/api/mappings/excel1`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selectedSheet: 'Consolidated_Data',
      fields: [
        { canonicalField: 's_no', mappedHeader: 'S.No', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'job_number', mappedHeader: 'Job No', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'balance_qty', mappedHeader: 'Balance Quantity', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'ocs_date', mappedHeader: 'OCS Date', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'head_rpo_to_job_monitor', mappedHeader: 'Head RPO to Job Monitor', confidence: 'HIGH', matchReason: 'Manual test' },
      ],
    }),
  });
  json = await checkJsonResponse(res, 'POST /api/mappings/excel1');
  assert(json.success === true && json.data.excel1.fields.length === 5, 'POST /api/mappings/excel1 updated dynamic mapping');

  // 1.6 POST /api/mappings/excel2
  res = await fetch(`${BASE_URL}/api/mappings/excel2`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selectedSheet: 'Inspection_Calls',
      fields: [
        { canonicalField: 'job_number', mappedHeader: 'Job No', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'inspection_from', mappedHeader: 'Inspection Date Attended From', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'inspection_upto', mappedHeader: 'Inspection Date Attended Upto', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'no_of_days', mappedHeader: 'No. of Days', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'qap_applicable', mappedHeader: 'QAP Approval', confidence: 'HIGH', matchReason: 'Manual test' },
      ],
    }),
  });
  json = await checkJsonResponse(res, 'POST /api/mappings/excel2');
  assert(json.success === true, 'POST /api/mappings/excel2 updated dynamic mapping');

  // 1.7 POST /api/mappings/excel3
  res = await fetch(`${BASE_URL}/api/mappings/excel3`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selectedSheet: 'Manhour_Audit_Master',
      fields: [
        { canonicalField: 'job_number', mappedHeader: 'Job No', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'running_orders', mappedHeader: 'No. of Running orders', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'orders_fd', mappedHeader: 'Orders for FD f/up', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'ocs_done', mappedHeader: 'OCS done', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'expediting', mappedHeader: 'Exp.', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'inspection', mappedHeader: 'Inspn', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'others', mappedHeader: 'Others', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'meeting', mappedHeader: 'Meeting', confidence: 'HIGH', matchReason: 'Manual test' },
        { canonicalField: 'total', mappedHeader: 'Total', confidence: 'HIGH', matchReason: 'Manual test' },
      ],
    }),
  });
  json = await checkJsonResponse(res, 'POST /api/mappings/excel3');
  assert(json.success === true, 'POST /api/mappings/excel3 updated dynamic mapping');

  // 1.8 404 test: Non-existent /api/random-route
  res = await fetch(`${BASE_URL}/api/some-nonexistent-endpoint-404`);
  json = await checkJsonResponse(res, 'GET /api/some-nonexistent-endpoint-404');
  assert(res.status === 404 && json.success === false, 'Non-existent API endpoint returns 404 JSON, NOT index.html');

  // ----------------------------------------------------
  // TEST SECTION 2: Replace File Verification (Excel 1, 2, 3)
  // ----------------------------------------------------
  console.log('\n--- SECTION 2: Replace File & Invalidation Verification ---');

  // Helper to generate a valid replacement workbook in memory
  function createTestWorkbook(sheetName: string, headers: string[], rows: any[][]) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  }

  // 2.1 Replace Excel 1 directly
  const buf1 = createTestWorkbook('Replaced_Data_1', ['Job No', 'Balance Quantity', 'OCS Date', 'Head RPO to Job Monitor'], [
    ['J-NEW-01', 2, '2026-08-10', ''],
  ]);
  const form1 = new FormData();
  form1.append('file', new Blob([buf1]), 'Replaced_Excel_1.xlsx');
  res = await fetch(`${BASE_URL}/api/workbooks/upload/excel1`, { method: 'POST', body: form1 });
  json = await checkJsonResponse(res, 'POST /api/workbooks/upload/excel1 (Direct Replacement)');
  assert(json.success === true && json.data.fileName === 'Replaced_Excel_1.xlsx', 'Excel 1 replaced directly');

  // Verify downstream invalidation: calculations and reviews should be empty
  res = await fetch(`${BASE_URL}/api/analysis/results`);
  json = await checkJsonResponse(res, 'GET /api/analysis/results after replacing Excel 1');
  assert(Object.keys(json.data || {}).length === 0, 'Downstream calculations were invalidated when Excel 1 was replaced');

  // Verify Excel 1 mapping was re-generated for new sheet
  res = await fetch(`${BASE_URL}/api/mappings`);
  json = await checkJsonResponse(res, 'GET /api/mappings after replacing Excel 1');
  assert(json.data.excel1.selectedSheet === 'Replaced_Data_1', 'Excel 1 mapping updated to new worksheet');

  // 2.2 Replace Excel 2 directly
  const buf2 = createTestWorkbook('Replaced_Calls_2', ['Job No', 'Inspection Date Attended From', 'Inspection Date Attended Upto', 'QAP Approval', 'No. of Days'], [
    ['J-NEW-01', '2026-08-15', '2026-08-15', '', 1],
  ]);
  const form2 = new FormData();
  form2.append('file', new Blob([buf2]), 'Replaced_Excel_2.xlsx');
  res = await fetch(`${BASE_URL}/api/workbooks/upload/excel2`, { method: 'POST', body: form2 });
  json = await checkJsonResponse(res, 'POST /api/workbooks/upload/excel2 (Direct Replacement)');
  assert(json.success === true && json.data.fileName === 'Replaced_Excel_2.xlsx', 'Excel 2 replaced directly');

  // Verify downstream invalidation again
  res = await fetch(`${BASE_URL}/api/analysis/results`);
  json = await checkJsonResponse(res, 'GET /api/analysis/results after replacing Excel 2');
  assert(Object.keys(json.data || {}).length === 0, 'Downstream calculations were invalidated when Excel 2 was replaced');

  // 2.3 Replace Excel 3 directly
  const buf3 = createTestWorkbook('Replaced_Template_3', ['Job No', 'No. of Running orders', 'Orders for FD f/up', 'OCS done', 'Exp.', 'Inspn', 'Others', 'Meeting', 'Total'], [
    ['J-NEW-01', 0, 0, 0, 0, 0, 0, 0, 0],
    ['Leave', 0, 0, 0, 0, 0, 0, 8, 8],
    ['Total', 0, 0, 0, 0, 0, 0, 8, 8],
  ]);
  const form3 = new FormData();
  form3.append('file', new Blob([buf3]), 'Replaced_Excel_3.xlsx');
  res = await fetch(`${BASE_URL}/api/workbooks/upload/excel3`, { method: 'POST', body: form3 });
  json = await checkJsonResponse(res, 'POST /api/workbooks/upload/excel3 (Direct Replacement)');
  assert(json.success === true && json.data.fileName === 'Replaced_Excel_3.xlsx', 'Excel 3 replaced directly');

  // Verify downstream invalidation again
  res = await fetch(`${BASE_URL}/api/analysis/results`);
  json = await checkJsonResponse(res, 'GET /api/analysis/results after replacing Excel 3');
  assert(Object.keys(json.data || {}).length === 0, 'Downstream calculations were invalidated when Excel 3 was replaced');

  // ----------------------------------------------------
  // TEST SECTION 3: Step 2 Dynamic Mapping & Validation
  // ----------------------------------------------------
  console.log('\n--- SECTION 3: Step 2 Validation & Proceed to Step 3 ---');

  // Reload standard samples so we have full standard sheets for reconciliation
  await fetch(`${BASE_URL}/api/workbooks/load-samples`, { method: 'POST' });

  // 3.1 Validate mapping initially with standard loaded samples
  res = await fetch(`${BASE_URL}/api/mappings/validate`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/mappings/validate');
  assert(json.success === true && json.data.isValid === true, 'Standard mapping validates successfully');

  // 3.2 Invalidate a required mapping in Excel 1 (set Job No to null)
  res = await fetch(`${BASE_URL}/api/mappings/excel1`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selectedSheet: 'Consolidated_Data',
      fields: [
        { canonicalField: 'job_number', mappedHeader: null, confidence: 'NONE', matchReason: 'Unmapped' },
        { canonicalField: 'balance_qty', mappedHeader: 'Balance Quantity', confidence: 'HIGH', matchReason: 'Test' },
      ],
    }),
  });
  json = await checkJsonResponse(res, 'POST /api/mappings/excel1 (Unmap required Job No)');
  assert(json.data.excel1.isValid === false, 'Excel 1 mapping marked invalid when required Job No is unmapped');

  // Re-run validate mappings to verify it fails
  res = await fetch(`${BASE_URL}/api/mappings/validate`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/mappings/validate with missing required field');
  assert(json.data.isValid === false, 'Validate Mapping correctly returns isValid: false when a required field is missing');
  assert(json.data.errors.length > 0, `Validation produced error messages: "${json.data.errors[0]}"`);

  // 3.3 Restore valid mapping
  res = await fetch(`${BASE_URL}/api/mappings/excel1`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selectedSheet: 'Consolidated_Data',
      fields: [
        { canonicalField: 's_no', mappedHeader: 'S.No', confidence: 'HIGH', matchReason: 'Restored' },
        { canonicalField: 'job_number', mappedHeader: 'Job No', confidence: 'HIGH', matchReason: 'Restored' },
        { canonicalField: 'balance_qty', mappedHeader: 'Balance Quantity', confidence: 'HIGH', matchReason: 'Restored' },
        { canonicalField: 'ocs_date', mappedHeader: 'OCS Date', confidence: 'HIGH', matchReason: 'Restored' },
        { canonicalField: 'head_rpo_to_job_monitor', mappedHeader: 'Head RPO to Job Monitor', confidence: 'HIGH', matchReason: 'Restored' },
      ],
    }),
  });
  json = await checkJsonResponse(res, 'POST /api/mappings/excel1 (Restore valid mapping)');
  if (!json.data.excel1.isValid) {
    console.error('DEBUG RESTORE ERRORS:', json.data.excel1.validationErrors);
  }
  assert(json.data.excel1.isValid === true, 'Excel 1 mapping restored to valid');

  res = await fetch(`${BASE_URL}/api/mappings/validate`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/mappings/validate (Re-validate after restoration)');
  assert(json.data.isValid === true, 'Validate Mapping confirms all workbooks are valid again');

  // 3.4 Proceed to Step 3: Trigger Job Reconciliation
  res = await fetch(`${BASE_URL}/api/jobs/reconcile`);
  json = await checkJsonResponse(res, 'GET /api/jobs/reconcile (Proceed to Step 3 action)');
  const reconciledJobsCount = json.data.totalJobs;
  assert(json.success === true && reconciledJobsCount > 0, `Job Reconciliation succeeded with ${reconciledJobsCount} jobs`);

  // Verify 'LEAVE' and 'TOTAL' rows were filtered out during reconciliation
  const allJobNumbers = (json.data.jobs || []).map((j: any) => j.jobNumber.toUpperCase());
  assert(!allJobNumbers.includes('LEAVE'), "'LEAVE' row correctly ignored from job list");
  assert(!allJobNumbers.includes('TOTAL'), "'TOTAL' row correctly ignored from job list");

  // ----------------------------------------------------
  // TEST SECTION 4: Version 2 Business Rules Verification
  // ----------------------------------------------------
  console.log('\n--- SECTION 4: Version 2 Business Rules Verification ---');

  res = await fetch(`${BASE_URL}/api/analysis/run`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ evaluationMonth: '2026-08' }),
  });
  json = await checkJsonResponse(res, 'POST /api/analysis/run (Step 4 Rule Analysis)');
  const calcs = json.data;
  
  // Job B269 Verification:
  // Excel 1:
  // Row 1: Bal=2, OCS=2026-08-10 => Running Order +1, OCS Done +1
  // Row 2: Bal=0, Head RPO blank => FD +1
  // Excel 2:
  // Row 1: From 2026-08-04 Upto 2026-08-05, Days=2, QAP blank => Inspn = 2 * 8 = 16h
  // Row 2: From blank, Upto blank, QAP Approved, Days=1 => Others = 1h
  // Formulas:
  // Expediting = 2 * (Running Orders + OCS Done) = 2 * (1 + 1) = 4h
  // Total = Exp (4) + Inspn (16) + Others (1) + Mtg (0) = 21h (FD is NOT included in Total)
  const b269 = calcs['B269'];
  assert(b269 !== undefined, 'Job B269 was calculated');
  if (b269) {
    assert(b269.runningOrders === 1, `B269 Running Orders = 1 (got: ${b269.runningOrders})`);
    assert(b269.ocsDone === 1, `B269 OCS Done = 1 (got: ${b269.ocsDone})`);
    assert(b269.fd === 1, `B269 FD = 1 (got: ${b269.fd})`);
    assert(b269.expediting === 4, `B269 Expediting = 2 * (1 + 1) = 4h (got: ${b269.expediting})`);
    assert(b269.inspection === 16, `B269 Inspection = 2 days * 8h = 16h (got: ${b269.inspection})`);
    assert(b269.others === 1, `B269 Others = 1h (got: ${b269.others})`);
    assert(b269.total === 21, `B269 Total = 4 + 16 + 1 = 21h (got: ${b269.total})`);
  }

  // Job B304 Verification:
  // Excel 1:
  // Row 1: Bal=5, OCS=2026-05-20 (outside Aug window 2026-07-26 to 2026-08-25) => Running Order +1, OCS Done +1
  // Row 2: Bal=3, OCS blank => Running Order +1, OCS Done 0
  // Running Orders = 2, OCS Done = 1 => Expediting = 2 * (2 + 1) = 6h
  // Excel 2:
  // Row 1: Days=3, Inspn = 3 * 8 = 24h
  // Row 2: Days=0.5, QAP Approved => Others = 0.5h (decimal preserved)
  // Total = 6 + 24 + 0.5 = 30.5h
  const b304 = calcs['B304'];
  assert(b304 !== undefined, 'Job B304 was calculated');
  if (b304) {
    assert(b304.runningOrders === 2, `B304 Running Orders = 2 (got: ${b304.runningOrders})`);
    assert(b304.ocsDone === 1, `B304 OCS Done = 1 (got: ${b304.ocsDone})`);
    assert(b304.expediting === 6, `B304 Expediting = 2 * (2 + 1) = 6h (got: ${b304.expediting})`);
    assert(b304.inspection === 24, `B304 Inspection = 3 days * 8h = 24h (got: ${b304.inspection})`);
    assert(b304.others === 0.5, `B304 Others = 0.5h decimal preserved (got: ${b304.others})`);
    assert(b304.total === 30.5, `B304 Total = 6 + 24 + 0.5 = 30.5h (got: ${b304.total})`);
  }

  // ----------------------------------------------------
  // TEST SECTION 5: Review, Overrides & Output Export
  // ----------------------------------------------------
  console.log('\n--- SECTION 5: Review, Overrides & Output Export ---');

  // Step 5: Auditor Review & Approve All
  res = await fetch(`${BASE_URL}/api/review/approve-all`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/review/approve-all (Step 5 Approval)');
  assert(json.success === true, 'Approved all jobs for output generation');

  // Step 6: Generate Output Workbook
  res = await fetch(`${BASE_URL}/api/output/generate`, { method: 'POST' });
  json = await checkJsonResponse(res, 'POST /api/output/generate (Step 6 Generation)');
  assert(json.success === true && json.data.totalJobsUpdated > 0, `Output generated with ${json.data.totalJobsUpdated} jobs updated`);

  // Verify checks in report
  const reportChecks = json.data.checks || [];
  const leavePreservedCheck = reportChecks.find((c: any) => c.name.includes('Leave & Total'));
  assert(leavePreservedCheck !== undefined, 'Validation report verifies Leave & Total template rows intact');

  // Step 6 Download Check
  res = await fetch(`${BASE_URL}/api/output/download`);
  const outContentType = res.headers.get('content-type') || '';
  assert(
    res.status === 200 && outContentType.includes('spreadsheetml'),
    `Output download returned HTTP 200 with Excel MIME type (${outContentType})`
  );

  console.log('\n====================================================');
  console.log(`VERIFICATION SUMMARY: ${passedCount} PASSED, ${failedCount} FAILED`);
  console.log('====================================================');

  if (failedCount > 0) {
    process.exit(1);
  }
}

runTests().catch((e) => {
  console.error('Fatal test error:', e);
  process.exit(1);
});
