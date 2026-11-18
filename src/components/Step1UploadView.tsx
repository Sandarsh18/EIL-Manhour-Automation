import React, { useState } from 'react';
import {
  ArrowRight,
  ShieldCheck,
  CheckCircle2,
  FileSpreadsheet,
  Info,
  Layers,
} from 'lucide-react';
import { WorkbookMetadata, WorkbookType, WorkbooksState } from '../types/eil.ts';
import { WorkbookUploadCard } from './WorkbookUploadCard.tsx';
import { QuickPreviewModal } from './QuickPreviewModal.tsx';

interface Step1UploadViewProps {
  workbooks: WorkbooksState;
  onUpload: (type: WorkbookType, file: File) => Promise<void>;
  onDelete: (type: WorkbookType) => Promise<void>;
  onProceedToStep2: () => void;
}

export const Step1UploadView: React.FC<Step1UploadViewProps> = ({
  workbooks,
  onUpload,
  onDelete,
  onProceedToStep2,
}) => {
  const [previewMeta, setPreviewMeta] = useState<WorkbookMetadata | null>(null);

  const workbookCount = [workbooks.excel1, workbooks.excel2, workbooks.excel3].filter(Boolean).length;
  const allUploaded = workbookCount === 3;
  const hasBlockingErrors = [workbooks.excel1, workbooks.excel2, workbooks.excel3].some(
    (wb) => wb?.validationStatus === 'ERROR'
  );
  const canProceed = allUploaded && !hasBlockingErrors;

  return (
    <div className="space-y-6">
      {/* Top Banner: Workflow Context */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white rounded-2xl p-6 shadow-md border border-slate-700/60">
        <div className="space-y-2 max-w-3xl">
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/20 text-blue-300 border border-blue-400/30">
              Step 1 of 6
            </span>
            <span className="text-xs text-slate-300 font-medium">Source Data Ingestion</span>
          </div>
          <h2 className="text-xl font-bold text-white tracking-tight">
            Upload Source Workbooks
          </h2>
          <p className="text-xs text-slate-300 leading-relaxed">
            Upload the three authoritative Excel workbooks independently. Each workbook undergoes structural
            validation, worksheet detection, and header inspection prior to dynamic column mapping in Step 2.
          </p>

          <div className="pt-1 flex flex-wrap items-center gap-4 text-xs text-slate-300">
            <div className="flex items-center gap-1.5">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              <span>Zero Hardcoded Positions</span>
            </div>
            <div className="flex items-center gap-1.5">
              <Layers className="w-4 h-4 text-blue-400" />
              <span>Isolated Server Processing</span>
            </div>
            <div className="flex items-center gap-1.5">
              <FileSpreadsheet className="w-4 h-4 text-indigo-400" />
              <span>Non-Destructive Master Output</span>
            </div>
          </div>
        </div>
      </div>

      {/* 3 Workbooks Ingestion Grid */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
        <WorkbookUploadCard
          type="excel1"
          metadata={workbooks.excel1}
          onUpload={onUpload}
          onDelete={() => onDelete('excel1')}
          onOpenPreview={(meta) => setPreviewMeta(meta)}
        />

        <WorkbookUploadCard
          type="excel2"
          metadata={workbooks.excel2}
          onUpload={onUpload}
          onDelete={() => onDelete('excel2')}
          onOpenPreview={(meta) => setPreviewMeta(meta)}
        />

        <WorkbookUploadCard
          type="excel3"
          metadata={workbooks.excel3}
          onUpload={onUpload}
          onDelete={() => onDelete('excel3')}
          onOpenPreview={(meta) => setPreviewMeta(meta)}
        />
      </div>

      {/* Audit & Business Rules Notice */}
      <div className="bg-blue-50/70 border border-blue-200 rounded-xl p-4 text-xs text-blue-900 flex items-start gap-3">
        <Info className="w-5 h-5 text-blue-600 shrink-0 mt-0.5" />
        <div className="space-y-1">
          <span className="font-bold text-blue-950">EIL Data Integrity & Rule Directives:</span>
          <p className="text-blue-800 leading-relaxed text-[11px]">
            • Columns are <strong>not assumed</strong> to be in fixed physical column indices. Dynamic mapping will occur in Step 2.
            <br />
            • OCS date evaluation evaluates the <strong>calendar month window</strong> (26th of previous month to 25th of evaluation month).
            <br />
            • Number of Days is <strong>authoritative</strong> for Inspection manhours (working days × 8).
            <br />
            • QAP Appl. priority rules strictly preserve <strong>numeric zeroes and decimals</strong> (0.25, 0.5, etc.).
            <br />
            • Excel 3 master template is <strong>never destructively modified</strong>; a distinct output workbook is produced.
          </p>
        </div>
      </div>

      {/* Bottom Readiness Bar & Proceed Action */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div
            className={`w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shrink-0 ${
              allUploaded
                ? 'bg-emerald-100 text-emerald-800'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            {workbookCount}/3
          </div>
          <div>
            <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <span>
                {allUploaded
                  ? 'All 3 Source Workbooks Ingested'
                  : `${3 - workbookCount} Workbook${3 - workbookCount > 1 ? 's' : ''} Remaining`}
              </span>
              {canProceed && (
                <CheckCircle2 className="w-4 h-4 text-emerald-600 inline" />
              )}
            </div>
            <div className="text-[11px] text-slate-500">
              {canProceed
                ? 'Workbook parsing and validation passed. Ready for dynamic column mapping.'
                : hasBlockingErrors
                ? 'One or more workbooks failed validation. Please resolve errors before continuing.'
                : 'Upload all three required workbooks to unlock Step 2.'}
            </div>
          </div>
        </div>

        <button
          id="proceed-step-2-btn"
          type="button"
          onClick={onProceedToStep2}
          disabled={!canProceed}
          className={`inline-flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs font-bold shadow-xs transition ${
            canProceed
              ? 'bg-blue-600 hover:bg-blue-700 text-white cursor-pointer'
              : 'bg-slate-200 text-slate-400 cursor-not-allowed'
          }`}
        >
          <span>Proceed to Step 2: Inspect & Map Columns</span>
          <ArrowRight className="w-4 h-4" />
        </button>
      </div>

      {/* Quick Preview Modal */}
      {previewMeta && (
        <QuickPreviewModal metadata={previewMeta} onClose={() => setPreviewMeta(null)} />
      )}
    </div>
  );
};
