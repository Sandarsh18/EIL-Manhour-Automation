/**
 * Frontend TypeScript contracts for EIL Manhour Automation & Audit
 */

export type WorkbookType = 'excel1' | 'excel2' | 'excel3';

export interface SheetSummary {
  name: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  previewRows: Record<string, any>[];
}

export type ValidationStatus = 'PENDING' | 'VALID' | 'WARNING' | 'ERROR';

export interface WorkbookMetadata {
  type: WorkbookType;
  fileName: string;
  fileSizeBytes: number;
  uploadedAt: string;
  sheets: SheetSummary[];
  activeSheet: string;
  validationStatus: ValidationStatus;
  validationErrors: string[];
  validationWarnings: string[];
}

export interface WorkbooksState {
  excel1: WorkbookMetadata | null;
  excel2: WorkbookMetadata | null;
  excel3: WorkbookMetadata | null;
}

export type MatchConfidence = 'HIGH' | 'MEDIUM' | 'LOW' | 'MANUAL' | 'UNMAPPED';

export type MappingFieldCategory = 'SOURCE_INPUT' | 'OUTPUT_TARGET';

export interface ColumnMappingField {
  canonicalField: string;
  label: string;
  description: string;
  required: boolean;
  mappedHeader: string | null;
  confidence: MatchConfidence;
  matchReason?: string;
  fieldCategory?: MappingFieldCategory;
  calculationSource?: string;
}

/**
 * Explicit contract for Excel 1 (Consolidated Report) Source Inputs
 */
export interface Excel1SourceMapping {
  s_no: string | null;
  job_number: string | null;
  foi_number?: string | null;
  balance_qty: string | null;
  ocs_date: string | null;
  head_rpo_to_job_monitor: string | null;
}

/**
 * Explicit contract for Excel 2 (Inspection Call Log) Source Inputs
 */
export interface Excel2SourceMapping {
  s_no: string | null;
  job_number: string | null;
  foi_number?: string | null;
  inspection_from: string | null;
  inspection_upto: string | null;
  review_dates?: string | null;
  qap_approval: string | null;
  number_of_days: string | null;
  contractor?: string | null;
  supplier?: string | null;
  item?: string | null;
}

/**
 * Explicit contract for Excel 3 (Master Template) Output Targets
 */
export interface Excel3OutputMapping {
  job_number: string | null;
  foi_number?: string | null;
  running_orders: string | null;
  orders_fd: string | null;
  ocs_done: string | null;
  expediting: string | null;
  inspection: string | null;
  others: string | null;
  total: string | null;
  meeting: string | null;
  remarks: string | null;
}

export interface WorkbookMappingConfig {
  workbookType: WorkbookType;
  selectedSheet: string;
  fields: ColumnMappingField[];
  isValid: boolean;
  validationErrors: string[];
  validationWarnings: string[];
}

export type MappingsState = Record<WorkbookType, WorkbookMappingConfig>;

export interface ReconciledJob {
  jobNumber: string;
  foiNumber?: string;
  inExcel1: boolean;
  inExcel2: boolean;
  inExcel3: boolean;
  excel1RowCount: number;
  excel2RowCount: number;
  excel2EligibleRowCount?: number;
  origin: 'Consolidated Report' | 'PM Call Logs' | 'Both';
  excel2EligibilityStatus: 'Eligible in Month' | 'No Eligible Calls' | 'N/A';
  matchStatus: 'MATCHED_ALL' | 'MISSING_EXCEL1' | 'MISSING_EXCEL2' | 'MISSING_EXCEL3' | 'EXCEL3_ONLY';
  selected: boolean;
}

export interface JobReconciliationSummary {
  totalJobs: number;
  matchedAll: number;
  excel1Count: number;
  excel2OnlyCount: number;
  missingExcel1: number;
  missingExcel2: number;
  missingExcel3: number;
  excel3Only: number;
  jobs: ReconciledJob[];
}

export interface Excel1OrderRecord {
  rowIndex: number;
  sNo: string;
  foiNumber?: string;
  poNumber?: string;
  orderNo: string; // Display Identifier (FOI No or PO No)
  description: string; // PO Description / Equipment
  supplier: string; // Supplier Name
  orderedQty?: number | string;
  balanceQty: number | string;
  headRpoDate?: string;
  ocsDate?: string;
  classification: 'RUNNING_ORDER' | 'OCS_DONE' | 'ORDERS_FD' | 'EXCLUDED_FD' | 'INVALID' | string;
  classificationLabel: string;
  classificationReason: string;
  expeditingMhContribution: number;
  rawRecord: Record<string, any>;
  sourceCell?: string;
  sourceSheet?: string;
  eligibility?: 'Eligible' | 'Excluded';
}

export interface Excel2InspectionRecord {
  rowIndex: number;
  callNo: string;
  foiNumber?: string;
  vendor: string;
  contractor: string;
  itemDescription: string;
  callStage: string;
  inspectionFrom?: string;
  inspectionUpto?: string;
  reviewDates?: string;
  qapApproval?: string;
  numberOfDays: number;
  multiplier: string;
  computedMh: number;
  status:
    | 'INCLUDED_INSPECTION'
    | 'INCLUDED_QAP'
    | 'INCLUDED_REVIEW'
    | 'INCLUDED_INSPECTION_WITH_CONFLICT'
    | 'INCLUDED_QAP_WITH_CONFLICT'
    | 'EXCLUDED_DATES'
    | 'ZERO_DAYS'
    | 'AMBIGUOUS'
    | 'NO_ACTIVITY'
    | string;
  reason: string;
  rawRecord: Record<string, any>;
  sourceCell?: string;
  sourceSheet?: string;
  eligibility?: 'Eligible' | 'Excluded';
}

export interface WarningAuditDetail {
  warningType: string;
  field: string;
  templateValue: number | string | null;
  calculatedValue: number | string;
  difference: number | string | null;
  explanation: string;
  supportingSourceRecords: string[];
  sourceRow?: number;
  sourceCell?: string;
  sourceWorkbook?: string;
  sourceSheet?: string;
  sourceValue?: any;
  problem?: string;
  effect?: string;
  recommendedAction?: string;
}

export interface CalculationLineageItem {
  field: string;
  sourceWorkbook: string;
  sourceSheet: string;
  sourceRow: number;
  column?: string;
  sourceCell?: string;
  rawValues: Record<string, any>;
  ruleApplied: string;
  intermediateCalculation?: string;
  calculatedValue: number | string;
  notes?: string;
}

export interface JobCalculationResult {
  jobNumber: string;
  foiNumber?: string;
  foiConflict?: boolean;
  origin?: 'Consolidated Report' | 'PM Call Logs' | 'Both';
  fd: number;
  runningOrders: number;
  ocsDone: number;
  expediting: number;
  inspection: number;
  others: number;
  meeting: number;
  total: number;
  status: 'VALID' | 'WARNING' | 'BLOCKED';
  warnings: string[];
  errors: string[];
  lineage: CalculationLineageItem[];
  sourceOrdersExcel1?: Excel1OrderRecord[];
  sourceCallsExcel2?: Excel2InspectionRecord[];
  warningAuditDetails?: WarningAuditDetail[];
}

export interface ReviewItem {
  jobNumber: string;
  foiNumber?: string;
  foiConflict?: boolean;
  origin?: 'Consolidated Report' | 'PM Call Logs' | 'Both';
  fd: number;
  runningOrders: number;
  ocsDone: number;
  expediting: number;
  inspection: number;
  others: number;
  meeting: number;
  total: number;
  approvalStatus: 'PENDING' | 'APPROVED' | 'BLOCKED' | 'FLAGGED';
  warnings: string[];
  errors: string[];
  isOverridden: boolean;
  overriddenFields: Record<string, number>;
  auditNotes: string;
  lineage: CalculationLineageItem[];
  sourceOrdersExcel1?: Excel1OrderRecord[];
  sourceCallsExcel2?: Excel2InspectionRecord[];
  warningAuditDetails?: WarningAuditDetail[];
}

export interface OutputFileMetadata {
  fileName: string;
  artifactId: string;
  downloadUrl: string;
  fileSizeBytes?: number;
  generatedAt?: string;
}

export interface ReportMetadata {
  fileName: string;
  generatedAt: string;
  evaluationMonth: string;
  totalJobs: number;
  populatedJobs: number;
  unmatchedTemplateRows: number;
  manualOverrides: number;
  totalManhours: number;
  status?: string;
}

export interface GeneratedOutputMetadata {
  fileName: string;
  pathOrArtifactId: string;
  generatedAt: string;
  jobCount: number;
  totalManhours: number;
  verifiedSampleJobs?: Array<{
    jobNumber: string;
    expediting: number;
    inspection: number;
    others: number;
    total: number;
    status: string;
    rawAddress?: string;
  }>;
}

export interface OutputValidationCheck {
  name: string;
  status: 'PASS' | 'WARN' | 'FAIL';
  details: string;
}

export interface OutputValidationReport {
  status: 'PASS' | 'FAIL' | 'WARNING';
  checks: OutputValidationCheck[];
  totalJobsUpdated: number;
  unmatchedJobsCount: number;
  errors: string[];
  warnings: string[];
  totalJobsToExport?: number;
  approvedCount?: number;
  skippedCount?: number;
  templateFileName?: string;
  generatedAt?: string;
  evaluationMonth?: string;
  outputDownloadUrl?: string;
  generatedOutput?: GeneratedOutputMetadata;
  reportMetadata?: ReportMetadata;
  outputFile?: OutputFileMetadata;
  verifiedSampleJobs?: Array<{
    jobNumber: string;
    targetRow?: number;
    allMatch?: boolean;
    cellValuesReadback: {
      fd: number;
      runningOrders: number;
      ocsDone: number;
      expediting: number;
      inspection: number;
      others: number;
      meeting: number;
      total: number;
    };
  }>;
}

export interface SavedProject {
  id: string;
  name: string;
  description?: string;
  createdAt: string;
  updatedAt: string;
  currentStep: number;
  evaluationMonth: string; // YYYY-MM
  workbooksState: WorkbooksState;
  mappingsState?: MappingsState;
  selectedJobs?: string[];
  calculations?: Record<string, JobCalculationResult>;
  reviewItems?: Record<string, ReviewItem>;
  outputReport?: OutputValidationReport | null;
}
