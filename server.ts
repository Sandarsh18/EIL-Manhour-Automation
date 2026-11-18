import express from 'express';
import path from 'path';
import multer from 'multer';
import { createServer as createViteServer } from 'vite';
import { ProjectStorage } from './server/storage.ts';
import { WorkbookType } from './server/types.ts';
import { EilDateParser } from './server/dateParser.ts';

const app = express();
const PORT = 3000;

// Multer in-memory storage for handling Excel uploads (max 50MB)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls' || ext === '.xlsm') {
      cb(null, true);
    } else {
      cb(new Error('Only Excel workbooks (.xlsx, .xls, .xlsm) are accepted'));
    }
  },
});

app.use(express.json());

// Initialize persistent storage directory
ProjectStorage.init();

// =================================================================
// API ROUTES (Mounted strictly BEFORE Vite / static fallback)
// =================================================================

// 1. Health check
app.get('/api/health', (_req, res) => {
  res.json({
    status: 'ok',
    service: 'EIL Manhour Automation & Audit Server',
    time: new Date().toISOString(),
  });
});

// 2. Workbooks endpoints
app.get('/api/workbooks', (_req, res) => {
  const state = ProjectStorage.getWorkbooksState();
  res.json({ success: true, data: state });
});

app.post('/api/workbooks/upload/:workbookType', (req, res) => {
  const workbookType = req.params.workbookType as WorkbookType;
  if (!['excel1', 'excel2', 'excel3'].includes(workbookType)) {
    return res.status(400).json({
      success: false,
      error: `Invalid workbook type: "${workbookType}". Must be excel1, excel2, or excel3.`,
    });
  }

  upload.single('file')(req, res, (uploadErr: any) => {
    if (uploadErr) {
      return res.status(400).json({
        success: false,
        error: uploadErr.message || 'Invalid file upload. Please select an Excel workbook (.xlsx, .xls, .xlsm).',
      });
    }

    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No file uploaded. Please select an Excel workbook (.xlsx, .xls).',
      });
    }

    try {
      const metadata = ProjectStorage.saveUploadedWorkbook(
        workbookType,
        req.file.originalname,
        req.file.buffer
      );
      return res.json({ success: true, data: metadata });
    } catch (err: any) {
      console.error(`Upload error for ${workbookType}:`, err);
      return res.status(500).json({
        success: false,
        error: `Failed to process ${workbookType}: ${err.message || err}`,
      });
    }
  });
});

app.delete('/api/workbooks/:workbookType', (req, res) => {
  const workbookType = req.params.workbookType as WorkbookType;
  if (!['excel1', 'excel2', 'excel3'].includes(workbookType)) {
    return res.status(400).json({ success: false, error: 'Invalid workbook type' });
  }
  ProjectStorage.deleteWorkbook(workbookType);
  return res.json({ success: true, message: `${workbookType} cleared` });
});

app.post('/api/workbooks/reset', (_req, res) => {
  ProjectStorage.resetAllWorkbooks();
  return res.json({
    success: true,
    data: ProjectStorage.getSessionState(),
    message: 'All active workbooks reset',
  });
});

// 3. Step 2 Column Mappings Endpoints
app.get('/api/mappings', (_req, res) => {
  const mappings = ProjectStorage.getMappingsState();
  res.json({ success: true, data: mappings });
});

app.post('/api/mappings/validate', (_req, res) => {
  const mappings = ProjectStorage.getMappingsState();
  const errors: string[] = [];
  const warnings: string[] = [];

  (['excel1', 'excel2', 'excel3'] as WorkbookType[]).forEach((type) => {
    const m = mappings[type];
    if (!m) {
      errors.push(`${type.toUpperCase()} mapping is missing.`);
    } else if (!m.isValid) {
      m.validationErrors.forEach((e) => errors.push(`${type.toUpperCase()}: ${e}`));
    }
  });

  res.json({
    success: true,
    data: {
      isValid: errors.length === 0,
      errors,
      warnings,
    },
  });
});

app.post('/api/mappings/:workbookType', (req, res) => {
  const workbookType = req.params.workbookType as WorkbookType;
  if (!['excel1', 'excel2', 'excel3'].includes(workbookType)) {
    return res.status(400).json({ success: false, error: 'Invalid workbook type' });
  }

  const { selectedSheet, fields } = req.body;
  if (!selectedSheet || !Array.isArray(fields)) {
    return res.status(400).json({
      success: false,
      error: 'selectedSheet and fields array are required.',
    });
  }

  try {
    const updated = ProjectStorage.updateMapping(workbookType, selectedSheet, fields);
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message || err });
  }
});

// 4. Step 3 Job Reconciliation Endpoint
app.get('/api/jobs/reconcile', (req, res) => {
  try {
    const evalMonth = (req.query.evaluationMonth as string) || undefined;
    const summary = ProjectStorage.reconcileJobs(evalMonth);
    res.json({ success: true, data: summary });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message || err });
  }
});

// 5. Step 4 Rule Analysis Endpoint
app.post('/api/analysis/run', (req, res) => {
  const { selectedJobs, evaluationMonth } = req.body;
  if (evaluationMonth) {
    ProjectStorage.setEvaluationMonth(evaluationMonth);
  }

  try {
    const results = ProjectStorage.runRuleAnalysis(selectedJobs);
    res.json({ success: true, data: results });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message || err });
  }
});

app.get('/api/analysis/results', (_req, res) => {
  const calculations = ProjectStorage.getCalculationResults();
  res.json({ success: true, data: calculations });
});

// 6. Step 5 Review & Approval Endpoints
app.get('/api/review', (_req, res) => {
  const items = ProjectStorage.getReviewItems();
  res.json({ success: true, data: items });
});

app.post('/api/review/approve-all', (_req, res) => {
  try {
    const updated = ProjectStorage.approveAllReviews();
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message || err });
  }
});

app.post('/api/review/unapprove-all', (_req, res) => {
  try {
    const updated = ProjectStorage.unapproveAllReviews();
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message || err });
  }
});

app.post('/api/review/delete-all', (_req, res) => {
  try {
    const updated = ProjectStorage.deleteAllReviews();
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message || err });
  }
});

app.delete('/api/review/:jobNumber', (req, res) => {
  const jobNumber = req.params.jobNumber;
  try {
    const updated = ProjectStorage.deleteReviewItem(jobNumber);
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message || err });
  }
});

app.post('/api/review/:jobNumber/delete', (req, res) => {
  const jobNumber = req.params.jobNumber;
  try {
    const updated = ProjectStorage.deleteReviewItem(jobNumber);
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message || err });
  }
});

app.post('/api/review/:jobNumber', (req, res) => {
  const jobNumber = req.params.jobNumber;
  try {
    const updated = ProjectStorage.updateReviewItem(jobNumber, req.body);
    res.json({ success: true, data: updated });
  } catch (err: any) {
    res.status(400).json({ success: false, error: err.message || err });
  }
});

// 7. Step 6 Output Generation & Download Endpoints
app.post('/api/output/generate', async (_req, res) => {
  try {
    const result = await ProjectStorage.generateOutput();
    res.json({
      success: true,
      data: result.report,
      report: result.report,
      reportMetadata: result.reportMetadata,
      outputFile: result.outputFile,
      generatedOutput: result.generatedOutput,
      fileName: result.outputFile.fileName,
      downloadUrl: result.outputFile.downloadUrl,
      artifactId: result.outputFile.artifactId,
    });
  } catch (err: any) {
    console.error('Output generation error:', err);
    res.status(500).json({
      success: false,
      error: {
        code: 'OUTPUT_GENERATION_FAILED',
        message: err.message || String(err),
      },
    });
  }
});

app.get('/api/output/download', (_req, res) => {
  const buffer = ProjectStorage.getOutputBuffer();
  if (!buffer) {
    return res.status(404).json({
      success: false,
      error: 'Output workbook is not available because generation failed.',
    });
  }

  const outMeta = ProjectStorage.getGeneratedOutput();
  const filename = outMeta?.fileName || 'EIL_Manhour_Audit_Output.xlsx';

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.send(buffer);
});

// 8. Workflow Step State
app.get('/api/session/state', (_req, res) => {
  const state = ProjectStorage.getSessionState();
  res.json({ success: true, data: state });
});

app.post('/api/workflow/step', (req, res) => {
  const { step, evaluationMonth } = req.body;
  if (typeof step === 'number') {
    ProjectStorage.setWorkflowStep(step);
  }
  if (evaluationMonth) {
    ProjectStorage.setEvaluationMonth(evaluationMonth);
  }
  res.json({ success: true });
});

// 9. Project Persistence Endpoints
app.get('/api/projects', (_req, res) => {
  const projects = ProjectStorage.listProjects();
  res.json({ success: true, data: projects });
});

app.post('/api/projects/save', (req, res) => {
  const { name, description } = req.body;
  if (!name || typeof name !== 'string' || !name.trim()) {
    return res.status(400).json({ success: false, error: 'Project name is required' });
  }
  const project = ProjectStorage.saveProject(name.trim(), description);
  res.json({ success: true, data: project });
});

app.post('/api/projects/load/:id', (req, res) => {
  const project = ProjectStorage.loadProject(req.params.id);
  if (!project) {
    return res.status(404).json({ success: false, error: 'Project not found' });
  }
  res.json({ success: true, data: project });
});

app.delete('/api/projects/:id', (req, res) => {
  ProjectStorage.deleteProject(req.params.id);
  res.json({ success: true, message: 'Project deleted' });
});

// 10. Date parsing & OCS window check utility
app.post('/api/validate-date', (req, res) => {
  const { dateVal, evaluationMonth } = req.body;
  const parsed = EilDateParser.parse(dateVal);
  let inOcsWindow = false;
  let ocsWindow = null;

  if (evaluationMonth) {
    try {
      ocsWindow = EilDateParser.getOcsEligibilityWindow(evaluationMonth);
      inOcsWindow = EilDateParser.isDateInOcsWindow(dateVal, evaluationMonth);
    } catch (e: any) {
      return res.status(400).json({ success: false, error: e.message });
    }
  }

  res.json({
    success: true,
    data: {
      parsed,
      inOcsWindow,
      ocsWindow,
    },
  });
});

// =================================================================
// 11. CRITICAL: API 404 Catch-All & Global API Error Handler
// These guarantee that ANY unknown or errored /api/* request ALWAYS
// returns JSON and NEVER falls through to index.html / <!doctype>!
// =================================================================
app.all('/api/*', (req, res) => {
  res.status(404).json({
    success: false,
    error: `API endpoint not found: ${req.method} ${req.originalUrl}`,
    path: req.originalUrl,
  });
});

app.use('/api/*', (err: any, req: any, res: any, next: any) => {
  console.error('[API Server Error]', err);
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({
    success: false,
    error: err.message || 'An unexpected internal server error occurred',
    path: req.originalUrl,
  });
});

// =================================================================
// 12. VITE / STATIC CLIENT SERVING (Mounted AFTER API routes)
// =================================================================
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[EIL Server] Running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
