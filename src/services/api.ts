import {
  WorkbookMetadata,
  WorkbookType,
  WorkbooksState,
  SavedProject,
  MappingsState,
  JobReconciliationSummary,
  JobCalculationResult,
  ReviewItem,
  OutputValidationReport,
} from '../types/eil.ts';

interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

/**
 * Hardened Central Request Helper
 * Ensures Content-Type validation, structured error diagnostics, and clean typed JSON return.
 * Eliminates "Unexpected token '<'" errors.
 */
async function apiRequest<T>(url: string, options: RequestInit = {}): Promise<T> {
  const method = options.method || 'GET';
  let response: Response;

  try {
    response = await fetch(url, options);
  } catch (netErr: any) {
    throw new Error(
      `Network connection failure: Unable to reach the EIL API server at ${url}. ${netErr.message || ''}`
    );
  }

  const contentType = response.headers.get('content-type') || '';

  // Validate that the server returned JSON
  if (!contentType.includes('application/json')) {
    const rawText = await response.text();
    const snippet = rawText.slice(0, 160).replace(/\s+/g, ' ').trim();
    throw new Error(
      `API request failed: Endpoint ${method} ${url} returned ${contentType || 'non-JSON'} instead of JSON (HTTP ${response.status}). Preview: "${snippet}". This usually indicates an API routing or dev server configuration issue.`
    );
  }

  let json: ApiResponse<T>;
  try {
    json = await response.json();
  } catch (parseErr: any) {
    throw new Error(
      `Failed to parse server JSON response from ${method} ${url}: ${parseErr.message}`
    );
  }

  if (!response.ok || !json.success) {
    const errObj = (json as any)?.error;
    const msg =
      typeof errObj === 'object' && errObj !== null
        ? errObj.message || errObj.code || JSON.stringify(errObj)
        : json.error || `Server responded with status ${response.status} for ${method} ${url}`;
    throw new Error(msg);
  }

  if (json.data && typeof json.data === 'object' && !Array.isArray(json.data)) {
    return {
      ...(json as any),
      ...json.data,
    } as T;
  }

  return json.data as T;
}

export const EilApi = {
  // Step 1: Workbooks
  async fetchWorkbooks(): Promise<WorkbooksState> {
    return apiRequest<WorkbooksState>('/api/workbooks');
  },

  async uploadWorkbook(type: WorkbookType, file: File): Promise<WorkbookMetadata> {
    const formData = new FormData();
    formData.append('file', file);

    return apiRequest<WorkbookMetadata>(`/api/workbooks/upload/${type}`, {
      method: 'POST',
      body: formData,
    });
  },

  async deleteWorkbook(type: WorkbookType): Promise<void> {
    await apiRequest<{ message: string }>(`/api/workbooks/${type}`, { method: 'DELETE' });
  },

  async resetWorkbooks(): Promise<void> {
    await apiRequest<{ message: string }>('/api/workbooks/reset', { method: 'POST' });
  },

  async loadSampleWorkbooks(): Promise<WorkbooksState> {
    return apiRequest<WorkbooksState>('/api/workbooks/load-samples', { method: 'POST' });
  },

  // Step 2: Column Mappings
  async getMappings(): Promise<MappingsState> {
    return apiRequest<MappingsState>('/api/mappings');
  },

  async updateMapping(
    type: WorkbookType,
    selectedSheet: string,
    fields: any[]
  ): Promise<MappingsState> {
    return apiRequest<MappingsState>(`/api/mappings/${type}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ selectedSheet, fields }),
    });
  },

  async validateMappings(): Promise<{ isValid: boolean; errors: string[]; warnings: string[] }> {
    return apiRequest<{ isValid: boolean; errors: string[]; warnings: string[] }>(
      '/api/mappings/validate',
      { method: 'POST' }
    );
  },

  // Step 3: Job Reconciliation
  async reconcileJobs(): Promise<JobReconciliationSummary> {
    return apiRequest<JobReconciliationSummary>('/api/jobs/reconcile');
  },

  // Step 4: Rule Analysis
  async runRuleAnalysis(
    selectedJobs: string[],
    evaluationMonth: string
  ): Promise<Record<string, JobCalculationResult>> {
    return apiRequest<Record<string, JobCalculationResult>>('/api/analysis/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ selectedJobs, evaluationMonth }),
    });
  },

  // Step 5: Review & Approval
  async getReviewItems(): Promise<Record<string, ReviewItem>> {
    return apiRequest<Record<string, ReviewItem>>('/api/review');
  },

  async updateReviewItem(
    jobNumber: string,
    updates: Partial<ReviewItem>
  ): Promise<ReviewItem> {
    return apiRequest<ReviewItem>(`/api/review/${encodeURIComponent(jobNumber)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(updates),
    });
  },

  async approveAllReviews(): Promise<Record<string, ReviewItem>> {
    return apiRequest<Record<string, ReviewItem>>('/api/review/approve-all', {
      method: 'POST',
    });
  },

  async unapproveAllReviews(): Promise<Record<string, ReviewItem>> {
    return apiRequest<Record<string, ReviewItem>>('/api/review/unapprove-all', {
      method: 'POST',
    });
  },

  async deleteAllReviews(): Promise<Record<string, ReviewItem>> {
    return apiRequest<Record<string, ReviewItem>>('/api/review/delete-all', {
      method: 'POST',
    });
  },

  async deleteReviewItem(jobNumber: string): Promise<Record<string, ReviewItem>> {
    return apiRequest<Record<string, ReviewItem>>(`/api/review/${encodeURIComponent(jobNumber)}`, {
      method: 'DELETE',
    });
  },

  async getSessionState(): Promise<{
    workbooks: any;
    mappings: any;
    currentStep: number;
    evaluationMonth: string;
    selectedJobs: string[];
    calculations: Record<string, any>;
    reviewItems: Record<string, ReviewItem>;
    outputReport: OutputValidationReport | null;
  }> {
    return apiRequest<any>('/api/session/state');
  },

  // Step 6: Output Generation
  async generateOutput(): Promise<OutputValidationReport> {
    const res = await apiRequest<any>(
      '/api/output/generate',
      { method: 'POST' }
    );

    // Support report directly, in res.report, in res.data, or combined
    const report: OutputValidationReport =
      res?.report || res?.data || res;

    if (!report || (!report.status && !report.checks && !report.reportMetadata && typeof report.totalJobsUpdated !== 'number')) {
      throw new Error('Output generation failed to return report metadata.');
    }

    if (res.reportMetadata && !report.reportMetadata) {
      report.reportMetadata = res.reportMetadata;
    }
    if (res.outputFile && !report.outputFile) {
      report.outputFile = res.outputFile;
    }
    if (res.generatedOutput && !report.generatedOutput) {
      report.generatedOutput = res.generatedOutput;
    }
    if (res.downloadUrl && !report.outputDownloadUrl) {
      report.outputDownloadUrl = res.downloadUrl;
    }

    return report;
  },

  async downloadOutputBlob(): Promise<{ blob: Blob; fileName: string }> {
    const res = await fetch(`/api/output/download?t=${Date.now()}`, {
      headers: {
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        Pragma: 'no-cache',
      },
    });

    if (!res.ok) {
      let errMsg = 'Output workbook is not available because generation failed.';
      try {
        const json = await res.json();
        if (json.error) {
          errMsg = typeof json.error === 'object' ? json.error.message || json.error.code : json.error;
        }
      } catch (_) {}
      throw new Error(errMsg);
    }

    const disposition = res.headers.get('Content-Disposition');
    let fileName = 'EIL_Manhour_Audit_Output.xlsx';
    if (disposition) {
      const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
      if (match && match[1]) {
        fileName = match[1].replace(/['"]/g, '').trim();
      }
    }

    const blob = await res.blob();
    if (blob.size === 0) {
      throw new Error('Downloaded file is empty. Generation must be executed first.');
    }

    return { blob, fileName };
  },

  getOutputDownloadUrl(): string {
    return `/api/output/download?t=${Date.now()}`;
  },

  // Workflow State & Evaluation Month
  async syncWorkflowStep(step: number, evaluationMonth?: string): Promise<void> {
    await apiRequest('/api/workflow/step', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ step, evaluationMonth }),
    });
  },

  // Projects
  async listProjects(): Promise<SavedProject[]> {
    return apiRequest<SavedProject[]>('/api/projects');
  },

  async saveProject(name: string, description?: string): Promise<SavedProject> {
    return apiRequest<SavedProject>('/api/projects/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, description }),
    });
  },

  async loadProject(id: string): Promise<SavedProject> {
    return apiRequest<SavedProject>(`/api/projects/load/${id}`, { method: 'POST' });
  },

  async deleteProject(id: string): Promise<void> {
    await apiRequest<{ message: string }>(`/api/projects/${id}`, { method: 'DELETE' });
  },

  // Date Check
  async validateDate(dateVal: any, evaluationMonth: string): Promise<any> {
    return apiRequest('/api/validate-date', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dateVal, evaluationMonth }),
    });
  },
};
