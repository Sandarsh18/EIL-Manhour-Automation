import fs from 'fs';
import path from 'path';
import {
  WorkbookMetadata,
  WorkbookType,
  WorkbooksState,
  SavedProject,
  MappingsState,
  JobCalculationResult,
  ReviewItem,
  OutputValidationReport,
  JobReconciliationSummary,
  GeneratedOutputMetadata,
  OutputFileMetadata,
  ReportMetadata,
} from './types.ts';
import { ExcelService, normalizeJobNumber, isValidJobNumber } from './excelService.ts';

const DATA_DIR = path.resolve(process.cwd(), 'data');
const UPLOADS_DIR = path.join(DATA_DIR, 'uploads');
const PROJECTS_FILE = path.join(DATA_DIR, 'projects.json');
const CURRENT_STATE_FILE = path.join(DATA_DIR, 'current_state.json');

export class ProjectStorage {
  private static workbooksState: WorkbooksState = {
    excel1: null,
    excel2: null,
    excel3: null,
  };

  private static rawBuffers: Map<WorkbookType, Buffer> = new Map();
  private static mappingsState: Partial<MappingsState> = {};
  private static currentStep: number = 1;
  private static evaluationMonth: string = '2026-08';
  private static selectedJobs: string[] = [];
  private static calculations: Record<string, JobCalculationResult> = {};
  private static reviewItems: Record<string, ReviewItem> = {};
  private static outputReport: OutputValidationReport | null = null;
  private static outputBuffer: Buffer | null = null;
  private static generatedOutput: GeneratedOutputMetadata | null = null;

  public static init() {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    if (!fs.existsSync(UPLOADS_DIR)) {
      fs.mkdirSync(UPLOADS_DIR, { recursive: true });
    }
    if (!fs.existsSync(PROJECTS_FILE)) {
      fs.writeFileSync(PROJECTS_FILE, JSON.stringify([], null, 2), 'utf-8');
    }

    // Restore current session state if available
    if (fs.existsSync(CURRENT_STATE_FILE)) {
      try {
        const raw = fs.readFileSync(CURRENT_STATE_FILE, 'utf-8');
        const parsed = JSON.parse(raw);
        if (parsed.workbooksState) this.workbooksState = parsed.workbooksState;
        if (parsed.mappingsState) this.mappingsState = parsed.mappingsState;
        if (parsed.currentStep) this.currentStep = parsed.currentStep;
        if (parsed.evaluationMonth) this.evaluationMonth = parsed.evaluationMonth;
        if (parsed.selectedJobs) this.selectedJobs = parsed.selectedJobs;
        if (parsed.calculations) {
          this.calculations = parsed.calculations;
          for (const calc of Object.values(this.calculations)) {
            if (calc && (calc.total === undefined || calc.total === null || isNaN(calc.total))) {
              calc.total = ExcelService.calculateTotal({
                expediting: calc.expediting || 0,
                inspection: calc.inspection || 0,
                others: calc.others || 0,
                meeting: calc.meeting || 0,
              });
            }
          }
        }
        if (parsed.reviewItems) {
          this.reviewItems = parsed.reviewItems;
          for (const item of Object.values(this.reviewItems)) {
            if (item && (item.total === undefined || item.total === null || isNaN(item.total))) {
              item.total = ExcelService.calculateTotal({
                expediting: item.expediting || 0,
                inspection: item.inspection || 0,
                others: item.others || 0,
                meeting: item.meeting || 0,
              });
            }
          }
        }
        if (parsed.outputReport) this.outputReport = parsed.outputReport;
        if (parsed.generatedOutput) this.generatedOutput = parsed.generatedOutput;

        // Restore buffers and re-parse metadata so row counts reflect accurate dimensions
        (['excel1', 'excel2', 'excel3'] as WorkbookType[]).forEach((type) => {
          try {
            const files = fs.readdirSync(UPLOADS_DIR);
            const matching = files.filter((f) => f.startsWith(`${type}_`)).sort();
            const found = matching[matching.length - 1];
            if (found) {
              const buf = fs.readFileSync(path.join(UPLOADS_DIR, found));
              this.rawBuffers.set(type, buf);
              const originalName = this.workbooksState[type]?.fileName || found.replace(new RegExp(`^${type}_[0-9]+_`), '');
              this.workbooksState[type] = ExcelService.parseWorkbookBuffer(buf, originalName, type);
              if ((!this.mappingsState[type] || !this.mappingsState[type]?.selectedSheet) && this.workbooksState[type]?.sheets.length) {
                const active = this.workbooksState[type]!.sheets.find((s) => s.name === this.workbooksState[type]!.activeSheet) || this.workbooksState[type]!.sheets[0];
                this.mappingsState[type] = ExcelService.generateDefaultMapping(type, active);
              }
            }
          } catch (e) {
            console.warn(`Could not restore buffer for ${type}:`, e);
          }
        });
      } catch (err) {
        console.warn('Failed to restore current state from disk:', err);
      }
    }
  }

  public static getWorkbooksState(): WorkbooksState {
    return this.workbooksState;
  }

  public static getWorkbookBuffer(type: WorkbookType): Buffer | null {
    if (this.rawBuffers.has(type)) {
      return this.rawBuffers.get(type)!;
    }
    const files = fs.readdirSync(UPLOADS_DIR);
    const found = files.find((f) => f.startsWith(`${type}_`));
    if (found) {
      const buf = fs.readFileSync(path.join(UPLOADS_DIR, found));
      this.rawBuffers.set(type, buf);
      return buf;
    }
    return null;
  }

  public static getMappingsState(): MappingsState {
    if (!this.mappingsState.excel1 || !this.mappingsState.excel1.selectedSheet) this.mappingsState.excel1 = this.createInitialMapping('excel1');
    if (!this.mappingsState.excel2 || !this.mappingsState.excel2.selectedSheet) this.mappingsState.excel2 = this.createInitialMapping('excel2');
    if (!this.mappingsState.excel3 || !this.mappingsState.excel3.selectedSheet) this.mappingsState.excel3 = this.createInitialMapping('excel3');
    return {
      excel1: this.mappingsState.excel1,
      excel2: this.mappingsState.excel2,
      excel3: this.mappingsState.excel3,
    };
  }

  public static updateMapping(type: WorkbookType, sheetName: string, fields: any[]): MappingsState {
    const wb = this.workbooksState[type];
    const sheet = wb?.sheets.find((s) => s.name === sheetName) || wb?.sheets[0];
    const headers = sheet?.headers || [];

    const { isValid, errors, warnings } = ExcelService.validateMappingConfig(type, sheetName, fields, headers);

    this.mappingsState[type] = {
      workbookType: type,
      selectedSheet: sheetName,
      fields,
      isValid,
      validationErrors: errors,
      validationWarnings: warnings,
    };

    // Invalidate downstream calculations when mapping changes
    this.calculations = {};
    this.reviewItems = {};
    this.outputReport = null;
    this.outputBuffer = null;

    this.persistCurrentState();
    return this.getMappingsState();
  }

  private static createInitialMapping(type: WorkbookType) {
    const wb = this.workbooksState[type];
    if (wb && wb.sheets.length > 0) {
      const activeSheet = wb.sheets.find((s) => s.name === wb.activeSheet) || wb.sheets[0];
      return ExcelService.generateDefaultMapping(type, activeSheet);
    }
    return {
      workbookType: type,
      selectedSheet: '',
      fields: [],
      isValid: false,
      validationErrors: ['Workbook not uploaded.'],
      validationWarnings: [],
    };
  }

  public static saveUploadedWorkbook(
    type: WorkbookType,
    originalName: string,
    buffer: Buffer
  ): WorkbookMetadata {
    this.init();

    // Clean up previous files for this type to prevent orphaned files
    try {
      const existingFiles = fs.readdirSync(UPLOADS_DIR);
      for (const f of existingFiles) {
        if (f.startsWith(`${type}_`)) {
          fs.unlinkSync(path.join(UPLOADS_DIR, f));
        }
      }
    } catch (e) {
      console.warn('Could not clean old upload:', e);
    }

    const fileName = `${type}_${Date.now()}_${originalName.replace(/[^a-zA-Z0-9._-]/g, '_')}`;
    const filePath = path.join(UPLOADS_DIR, fileName);

    fs.writeFileSync(filePath, buffer);
    this.rawBuffers.set(type, buffer);

    const metadata = ExcelService.parseWorkbookBuffer(buffer, originalName, type);
    this.workbooksState[type] = metadata;

    // Automatic default mapping generation using fuzzy matching
    if (metadata.sheets.length > 0) {
      const activeSheet = metadata.sheets.find((s) => s.name === metadata.activeSheet) || metadata.sheets[0];
      this.mappingsState[type] = ExcelService.generateDefaultMapping(type, activeSheet);
    } else {
      delete this.mappingsState[type];
    }

    // Invalidate downstream calculations, review, and output because source workbook changed
    this.calculations = {};
    this.reviewItems = {};
    this.outputReport = null;
    this.outputBuffer = null;
    this.selectedJobs = [];
    this.currentStep = 1;

    this.persistCurrentState();
    return metadata;
  }

  public static deleteWorkbook(type: WorkbookType): boolean {
    this.workbooksState[type] = null;
    this.rawBuffers.delete(type);
    delete this.mappingsState[type];

    try {
      const existingFiles = fs.readdirSync(UPLOADS_DIR);
      for (const f of existingFiles) {
        if (f.startsWith(`${type}_`)) {
          fs.unlinkSync(path.join(UPLOADS_DIR, f));
        }
      }
    } catch (e) {
      console.warn('Could not clean upload on delete:', e);
    }

    // Invalidate downstream calculations and step state
    this.calculations = {};
    this.reviewItems = {};
    this.outputReport = null;
    this.outputBuffer = null;
    this.currentStep = 1;

    this.persistCurrentState();
    return true;
  }

  public static resetAllWorkbooks(): void {
    this.workbooksState = { excel1: null, excel2: null, excel3: null };
    this.rawBuffers.clear();
    this.mappingsState = {};
    this.currentStep = 1;
    this.selectedJobs = [];
    this.calculations = {};
    this.reviewItems = {};
    this.outputReport = null;
    this.outputBuffer = null;

    try {
      const existingFiles = fs.readdirSync(UPLOADS_DIR);
      for (const f of existingFiles) {
        fs.unlinkSync(path.join(UPLOADS_DIR, f));
      }
    } catch (e) {
      console.warn('Could not clean uploads on reset:', e);
    }

    this.persistCurrentState();
  }

  public static loadSampleData(): WorkbooksState {
    this.init();
    const samples = ExcelService.createSampleWorkbooks();

    this.saveUploadedWorkbook('excel1', samples.excel1.fileName, samples.excel1.buffer);
    this.saveUploadedWorkbook('excel2', samples.excel2.fileName, samples.excel2.buffer);
    this.saveUploadedWorkbook('excel3', samples.excel3.fileName, samples.excel3.buffer);

    if (this.workbooksState.excel1) this.workbooksState.excel1.isSampleData = true;
    if (this.workbooksState.excel2) this.workbooksState.excel2.isSampleData = true;
    if (this.workbooksState.excel3) this.workbooksState.excel3.isSampleData = true;

    this.persistCurrentState();
    return this.workbooksState;
  }

  // Reconciliation
  public static reconcileJobs(evaluationMonth?: string): JobReconciliationSummary {
    const evalMonth = evaluationMonth || this.evaluationMonth;
    if (evaluationMonth) {
      this.evaluationMonth = evaluationMonth;
    }
    const buf1 = this.getWorkbookBuffer('excel1');
    const buf2 = this.getWorkbookBuffer('excel2');
    const buf3 = this.getWorkbookBuffer('excel3');
    const mappings = this.getMappingsState();

    if (!buf1 || !buf2 || !buf3) {
      throw new Error('All three workbooks must be uploaded to perform job reconciliation.');
    }

    const summary = ExcelService.reconcileJobs(
      buf1,
      mappings.excel1,
      buf2,
      mappings.excel2,
      buf3,
      mappings.excel3,
      evalMonth
    );

    // Update selectedJobs default if empty: jobs in authoritative universe J1 UNION J2
    if (this.selectedJobs.length === 0) {
      this.selectedJobs = summary.jobs.filter((j) => j.selected).map((j) => j.jobNumber);
    }

    this.persistCurrentState();
    return summary;
  }

  // Rule Analysis
  public static runRuleAnalysis(selectedJobNumbers?: string[]): Record<string, JobCalculationResult> {
    const buf1 = this.getWorkbookBuffer('excel1');
    const buf2 = this.getWorkbookBuffer('excel2');
    const buf3 = this.getWorkbookBuffer('excel3');
    const mappings = this.getMappingsState();

    if (!buf1 || !buf2 || !buf3) {
      throw new Error('All three workbooks must be uploaded before running rule analysis.');
    }

    const summary = this.reconcileJobs(this.evaluationMonth);
    const validUniverseJobs = new Set(summary.jobs.map((j) => j.jobNumber));

    if (selectedJobNumbers && selectedJobNumbers.length > 0) {
      this.selectedJobs = selectedJobNumbers
        .filter((j) => j && typeof j === 'string' && isValidJobNumber(j))
        .map((j) => normalizeJobNumber(j))
        .filter((j) => validUniverseJobs.has(j));
    } else {
      this.selectedJobs = this.selectedJobs
        .filter((j) => j && typeof j === 'string' && isValidJobNumber(j))
        .map((j) => normalizeJobNumber(j))
        .filter((j) => validUniverseJobs.has(j));
    }

    if (this.selectedJobs.length === 0) {
      this.selectedJobs = summary.jobs.filter((j) => j.selected).map((j) => j.jobNumber).filter((j) => isValidJobNumber(j));
    }

    const data1 = ExcelService.extractSheetRows(buf1, mappings.excel1.selectedSheet);
    const data2 = ExcelService.extractSheetRows(buf2, mappings.excel2.selectedSheet);
    const data3 = ExcelService.extractSheetRows(buf3, mappings.excel3.selectedSheet);

    const results: Record<string, JobCalculationResult> = {};
    const reviewItems: Record<string, ReviewItem> = {};

    for (const jn of this.selectedJobs) {
      if (!jn || !isValidJobNumber(jn)) continue;
      const existingReview = this.reviewItems[jn];
      const existingMtg = existingReview?.overriddenFields?.meeting ?? existingReview?.meeting ?? 0;
      const calc = ExcelService.analyzeJob(
        jn,
        this.evaluationMonth,
        data1.rows,
        mappings.excel1,
        data2.rows,
        mappings.excel2,
        data3.rows,
        mappings.excel3,
        existingMtg,
        data1.headers,
        data2.headers,
        data3.headers
      );
      results[jn] = calc;

      // Initialize review item
      reviewItems[jn] = {
        jobNumber: calc.jobNumber,
        foiNumber: calc.foiNumber,
        foiConflict: calc.foiConflict,
        origin: calc.origin,
        fd: calc.fd,
        runningOrders: calc.runningOrders,
        ocsDone: calc.ocsDone,
        expediting: calc.expediting,
        inspection: calc.inspection,
        others: calc.others,
        meeting: calc.meeting,
        total: calc.total,
        approvalStatus: calc.status === 'BLOCKED' ? 'BLOCKED' : (existingReview?.approvalStatus || 'PENDING'),
        warnings: calc.warnings,
        errors: calc.errors,
        isOverridden: existingReview?.isOverridden || false,
        overriddenFields: existingReview?.overriddenFields || {},
        auditNotes: existingReview?.auditNotes || '',
        lineage: calc.lineage,
        sourceOrdersExcel1: calc.sourceOrdersExcel1,
        sourceCallsExcel2: calc.sourceCallsExcel2,
        warningAuditDetails: calc.warningAuditDetails,
      };
    }

    this.calculations = results;
    this.reviewItems = reviewItems;
    this.persistCurrentState();
    return results;
  }

  public static getCalculationResults(): Record<string, JobCalculationResult> {
    return this.calculations;
  }

  // Review & Approval
  public static getReviewItems(): Record<string, ReviewItem> {
    return this.reviewItems;
  }

  public static updateReviewItem(
    jobNumber: string,
    updates: Partial<ReviewItem>
  ): ReviewItem {
    const current = this.reviewItems[jobNumber];
    if (!current) throw new Error(`Job "${jobNumber}" not found in active review items.`);

    const updated = { ...current, ...updates };

    // Auto-update expediting if runningOrders or ocsDone changed and expediting not explicitly overridden
    if ((updates.runningOrders !== undefined || updates.ocsDone !== undefined) && updates.expediting === undefined && !updated.overriddenFields?.expediting) {
      updated.expediting = 2 * ((Number(updated.runningOrders) || 0) + (Number(updated.ocsDone) || 0));
    }

    // Recompute total per authoritative Version 2 rule: Total = Exp + Inspn + Others + Meeting
    updated.total = ExcelService.calculateTotal({
      expediting: updated.expediting,
      inspection: updated.inspection,
      others: updated.others,
      meeting: updated.meeting,
    });

    this.reviewItems[jobNumber] = updated;
    this.persistCurrentState();
    return updated;
  }

  public static approveAllReviews(): Record<string, ReviewItem> {
    for (const [jn, item] of Object.entries(this.reviewItems)) {
      if (item.approvalStatus !== 'BLOCKED') {
        item.approvalStatus = 'APPROVED';
      }
    }
    this.persistCurrentState();
    return this.reviewItems;
  }

  public static unapproveAllReviews(): Record<string, ReviewItem> {
    for (const [jn, item] of Object.entries(this.reviewItems)) {
      if (item.approvalStatus !== 'BLOCKED') {
        item.approvalStatus = 'PENDING';
      }
    }
    this.persistCurrentState();
    return this.reviewItems;
  }

  public static deleteAllReviews(): Record<string, ReviewItem> {
    this.reviewItems = {};
    this.calculations = {};
    this.selectedJobs = [];
    this.outputBuffer = null;
    this.outputReport = null;
    this.persistCurrentState();
    return this.reviewItems;
  }

  public static deleteReviewItem(jobNumber: string): Record<string, ReviewItem> {
    if (!jobNumber || typeof jobNumber !== 'string' || !jobNumber.trim()) {
      return this.reviewItems;
    }
    const rawTrimmed = jobNumber.trim();
    const norm = isValidJobNumber(rawTrimmed) ? normalizeJobNumber(rawTrimmed) : rawTrimmed.toUpperCase();

    delete this.reviewItems[jobNumber];
    delete this.reviewItems[rawTrimmed];
    delete this.reviewItems[norm];

    delete this.calculations[jobNumber];
    delete this.calculations[rawTrimmed];
    delete this.calculations[norm];

    this.selectedJobs = this.selectedJobs.filter((j) => {
      if (!j || !j.trim()) return false;
      const jTrimmed = j.trim();
      const jNorm = isValidJobNumber(jTrimmed) ? normalizeJobNumber(jTrimmed) : jTrimmed.toUpperCase();
      return j !== jobNumber && j !== rawTrimmed && j !== norm && jNorm !== norm && jTrimmed !== rawTrimmed;
    });

    this.outputBuffer = null;
    this.outputReport = null;
    this.generatedOutput = null;
    this.persistCurrentState();
    return this.reviewItems;
  }

  public static getSessionState() {
    return {
      workbooks: this.getWorkbooksState(),
      mappings: this.getMappingsState(),
      currentStep: this.currentStep,
      evaluationMonth: this.evaluationMonth,
      selectedJobs: this.selectedJobs,
      calculations: this.calculations,
      reviewItems: this.reviewItems,
      outputReport: this.outputReport,
    };
  }

  // Output Generation
  public static async generateOutput(): Promise<{
    report: OutputValidationReport;
    generatedOutput: GeneratedOutputMetadata;
    reportMetadata: ReportMetadata;
    outputFile: OutputFileMetadata;
  }> {
    const buf3 = this.getWorkbookBuffer('excel3');
    const mappings = this.getMappingsState();

    if (!buf3) {
      throw new Error('Excel 3 master template buffer is missing. Please upload the master template in Step 1.');
    }

    const { buffer, report } = await ExcelService.generateOutputWorkbook(
      buf3,
      mappings.excel3,
      this.reviewItems,
      this.evaluationMonth
    );

    this.outputBuffer = buffer;

    // Save output file to disk with clear output artifact naming
    const outFileName = `EIL_Manhour_Audit_Output_${this.evaluationMonth}_${Date.now()}.xlsx`;
    const outPath = path.join(UPLOADS_DIR, outFileName);
    fs.writeFileSync(outPath, buffer);

    const reviewItemsList = Object.values(this.reviewItems);
    const approvedList = reviewItemsList.filter((i) => i.approvalStatus === 'APPROVED');
    const totalManhours = approvedList.reduce((sum, i) => sum + (i.total || 0), 0);
    const manualOverrides = reviewItemsList.filter((i) => i.isOverridden).length;

    const outMeta: GeneratedOutputMetadata = {
      fileName: outFileName,
      pathOrArtifactId: outPath,
      generatedAt: new Date().toISOString(),
      jobCount: report.totalJobsUpdated || reviewItemsList.length,
      totalManhours: totalManhours,
      verifiedSampleJobs: report.generatedOutput?.verifiedSampleJobs || [],
    };

    const reportMeta: ReportMetadata = {
      fileName: outFileName,
      generatedAt: outMeta.generatedAt,
      evaluationMonth: this.evaluationMonth,
      totalJobs: reviewItemsList.length,
      populatedJobs: report.totalJobsUpdated || reviewItemsList.length,
      unmatchedTemplateRows: report.unmatchedJobsCount || 0,
      manualOverrides,
      totalManhours,
      status: report.status,
    };

    const fileMeta: OutputFileMetadata = {
      fileName: outFileName,
      artifactId: outPath,
      downloadUrl: '/api/output/download',
      fileSizeBytes: buffer.length,
      generatedAt: outMeta.generatedAt,
    };

    this.generatedOutput = outMeta;
    report.generatedOutput = outMeta;
    report.reportMetadata = reportMeta;
    report.outputFile = fileMeta;
    report.evaluationMonth = this.evaluationMonth;
    report.outputDownloadUrl = '/api/output/download';
    this.outputReport = report;

    this.persistCurrentState();
    return {
      report,
      generatedOutput: outMeta,
      reportMetadata: reportMeta,
      outputFile: fileMeta,
    };
  }

  public static getGeneratedOutput(): GeneratedOutputMetadata | null {
    return this.generatedOutput;
  }

  public static getOutputBuffer(): Buffer | null {
    if (this.outputBuffer) {
      return this.outputBuffer;
    }
    // Check disk for latest generated output file if in-memory buffer was cleared
    try {
      if (fs.existsSync(UPLOADS_DIR)) {
        const files = fs.readdirSync(UPLOADS_DIR)
          .filter((f) => f.startsWith('EIL_Manhour_Audit_Output_') && f.endsWith('.xlsx'))
          .map((f) => ({
            name: f,
            path: path.join(UPLOADS_DIR, f),
            mtime: fs.statSync(path.join(UPLOADS_DIR, f)).mtimeMs,
          }))
          .sort((a, b) => b.mtime - a.mtime);

        if (files.length > 0) {
          const latest = fs.readFileSync(files[0].path);
          this.outputBuffer = latest;
          if (!this.generatedOutput) {
            this.generatedOutput = {
              fileName: files[0].name,
              pathOrArtifactId: files[0].path,
              generatedAt: new Date(files[0].mtime).toISOString(),
              jobCount: Object.keys(this.reviewItems).length,
              totalManhours: 0,
              verifiedSampleJobs: this.outputReport?.generatedOutput?.verifiedSampleJobs || [],
            };
          }
          return latest;
        }
      }
    } catch (e) {
      console.warn('Error recovering output buffer from disk:', e);
    }
    return null;
  }

  // Session & Project Management
  public static setWorkflowStep(step: number) {
    this.currentStep = step;
    this.persistCurrentState();
  }

  public static setEvaluationMonth(month: string) {
    this.evaluationMonth = month;
    this.persistCurrentState();
  }

  public static listProjects(): SavedProject[] {
    this.init();
    try {
      const data = fs.readFileSync(PROJECTS_FILE, 'utf-8');
      return JSON.parse(data) as SavedProject[];
    } catch {
      return [];
    }
  }

  public static saveProject(name: string, description?: string): SavedProject {
    this.init();
    const projects = this.listProjects();
    const now = new Date().toISOString();

    const project: SavedProject = {
      id: `proj_${Date.now()}`,
      name,
      description: description || '',
      createdAt: now,
      updatedAt: now,
      currentStep: this.currentStep,
      evaluationMonth: this.evaluationMonth,
      workbooksState: { ...this.workbooksState },
      mappingsState: this.getMappingsState(),
      selectedJobs: [...this.selectedJobs],
      calculations: { ...this.calculations },
      reviewItems: { ...this.reviewItems },
      outputReport: this.outputReport,
    };

    projects.unshift(project);
    fs.writeFileSync(PROJECTS_FILE, JSON.stringify(projects, null, 2), 'utf-8');
    return project;
  }

  public static loadProject(projectId: string): SavedProject | null {
    this.init();
    const projects = this.listProjects();
    const project = projects.find((p) => p.id === projectId);
    if (!project) return null;

    this.workbooksState = { ...project.workbooksState };
    if (project.mappingsState) this.mappingsState = project.mappingsState;
    if (project.currentStep) this.currentStep = project.currentStep;
    if (project.evaluationMonth) this.evaluationMonth = project.evaluationMonth;
    if (project.selectedJobs) this.selectedJobs = project.selectedJobs;
    if (project.calculations) this.calculations = project.calculations;
    if (project.reviewItems) this.reviewItems = project.reviewItems;
    if (project.outputReport) this.outputReport = project.outputReport;

    this.persistCurrentState();
    return project;
  }

  public static deleteProject(projectId: string): boolean {
    this.init();
    const projects = this.listProjects();
    const filtered = projects.filter((p) => p.id !== projectId);
    fs.writeFileSync(PROJECTS_FILE, JSON.stringify(filtered, null, 2), 'utf-8');
    return true;
  }

  private static persistCurrentState() {
    try {
      const state = {
        workbooksState: this.workbooksState,
        mappingsState: this.mappingsState,
        currentStep: this.currentStep,
        evaluationMonth: this.evaluationMonth,
        selectedJobs: this.selectedJobs,
        calculations: this.calculations,
        reviewItems: this.reviewItems,
        outputReport: this.outputReport,
        generatedOutput: this.generatedOutput,
      };
      fs.writeFileSync(CURRENT_STATE_FILE, JSON.stringify(state, null, 2), 'utf-8');
    } catch (e) {
      console.error('Error persisting current state:', e);
    }
  }
}
