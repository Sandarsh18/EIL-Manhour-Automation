import React, { useRef, useState } from 'react';
import {
  Upload,
  FileSpreadsheet,
  AlertCircle,
  CheckCircle2,
  AlertTriangle,
  Trash2,
  Eye,
  RefreshCw,
  Loader2,
} from 'lucide-react';
import { WorkbookMetadata, WorkbookType } from '../types/eil.ts';

interface WorkbookUploadCardProps {
  type: WorkbookType;
  metadata: WorkbookMetadata | null;
  onUpload: (type: WorkbookType, file: File) => Promise<void>;
  onDelete: () => void;
  onOpenPreview: (metadata: WorkbookMetadata) => void;
}

const WORKBOOK_DESCRIPTIONS: Record<
  WorkbookType,
  {
    title: string;
    description: string;
    keyFields: string[];
    bgAccent: string;
  }
> = {
  excel1: {
    title: 'Excel 1 — CONSOLIDATED REPORT',
    description:
      'Authoritative source for FD, Balance Quantity, Running Orders, OCS Date eligibility and related Excel 1 calculations.',
    keyFields: ['Job Number', 'Balance Quantity', 'OCS Date', 'Status', 'PO Number'],
    bgAccent: 'border-l-blue-600',
  },
  excel2: {
    title: 'Excel 2 — INSPECTION CALL LOG',
    description:
      'Inspection call records, valid working days × 8h, QAP Appl., priority rules and Head RPO.',
    keyFields: [
      'Job Number',
      'Inspection Attended From',
      'Inspection Attended Upto',
      'No. of Days',
      'Head RPO',
      'QAP Appl.',
    ],
    bgAccent: 'border-l-teal-600',
  },
  excel3: {
    title: 'Excel 3 — MASTER TEMPLATE',
    description:
      'Baseline reporting structure containing the master job list and output calculation columns.',
    keyFields: [
      'Job Number',
      'Running Orders',
      'OCS Done',
      'Expediting',
      'Inspection',
      'Others',
      'Meeting',
      'Total',
    ],
    bgAccent: 'border-l-indigo-600',
  },
};

export const WorkbookUploadCard: React.FC<WorkbookUploadCardProps> = ({
  type,
  metadata,
  onUpload,
  onDelete,
  onOpenPreview,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [isReplacing, setIsReplacing] = useState(false);
  const [isConfirmingDelete, setIsConfirmingDelete] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const info = WORKBOOK_DESCRIPTIONS[type];

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      await processSelectedFile(file);
    }
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      await processSelectedFile(file);
    }
    // Reset file input so the same filename can be re-selected if needed
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const processSelectedFile = async (file: File) => {
    setUploadError(null);
    const ext = file.name.toLowerCase();
    if (!ext.endsWith('.xlsx') && !ext.endsWith('.xls') && !ext.endsWith('.xlsm')) {
      setUploadError('Invalid format. Please select an Excel workbook (.xlsx, .xls).');
      return;
    }

    try {
      setIsUploading(true);
      await onUpload(type, file);
    } catch (err: any) {
      setUploadError(err.message || 'Failed to upload and parse Excel workbook');
    } finally {
      setIsUploading(false);
      setIsReplacing(false);
    }
  };

  const handleTriggerReplace = () => {
    setIsReplacing(true);
    setUploadError(null);
    fileInputRef.current?.click();
  };

  const formatFileSize = (bytes?: number | null): string => {
    if (bytes == null || isNaN(Number(bytes))) return '0 B';
    const numBytes = Number(bytes);
    if (numBytes < 1024) return `${numBytes} B`;
    if (numBytes < 1024 * 1024) return `${(numBytes / 1024).toFixed(1)} KB`;
    return `${(numBytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const activeSheetData = metadata?.sheets.find((s) => s.name === metadata.activeSheet) || metadata?.sheets[0];

  return (
    <div
      id={`card-${type}`}
      className={`bg-white rounded-xl border border-slate-200 border-l-4 ${info.bgAccent} shadow-xs flex flex-col justify-between transition overflow-hidden`}
    >
      {/* Top Card Header */}
      <div className="p-4 border-b border-slate-100">
        <div className="flex items-start justify-between gap-2">
          <div>
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Source Workbook {type === 'excel1' ? '1' : type === 'excel2' ? '2' : '3'}
            </span>
            <h3 className="text-sm font-bold text-slate-900 leading-tight">{info.title}</h3>
          </div>

          {/* Status Badge */}
          <div>
            {isUploading ? (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-blue-50 text-blue-700 border border-blue-200">
                <Loader2 className="w-3 h-3 animate-spin" />
                {isReplacing ? 'Replacing...' : 'Uploading...'}
              </span>
            ) : !metadata ? (
              <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-medium bg-slate-100 text-slate-600 border border-slate-200">
                Pending Upload
              </span>
            ) : metadata.validationStatus === 'VALID' ? (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                <CheckCircle2 className="w-3 h-3" />
                Valid
              </span>
            ) : metadata.validationStatus === 'WARNING' ? (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-50 text-amber-700 border border-amber-200">
                <AlertTriangle className="w-3 h-3" />
                Warning
              </span>
            ) : (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-rose-50 text-rose-700 border border-rose-200">
                <AlertCircle className="w-3 h-3" />
                Error
              </span>
            )}
          </div>
        </div>

        <p className="text-xs text-slate-500 mt-2 leading-relaxed">{info.description}</p>
      </div>

      {/* Card Body: Dropzone or Uploaded Details */}
      <div className="p-4 flex-1 flex flex-col justify-between">
        {!metadata ? (
          /* Empty / Dropzone State */
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            className={`flex-1 border-2 border-dashed rounded-lg p-5 flex flex-col items-center justify-center text-center transition ${
              isDragging
                ? 'border-blue-500 bg-blue-50/50'
                : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.xlsm"
              className="hidden"
              onChange={handleFileChange}
            />

            {isUploading ? (
              <div className="py-4 flex flex-col items-center gap-2">
                <Loader2 className="w-7 h-7 text-blue-600 animate-spin" />
                <span className="text-xs font-medium text-slate-700">
                  Parsing workbook on server...
                </span>
              </div>
            ) : (
              <>
                <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-500 mb-2">
                  <Upload className="w-5 h-5" />
                </div>
                <div className="text-xs font-semibold text-slate-800">
                  Drag & drop your Excel file here
                </div>
                <div className="text-[11px] text-slate-500 mt-1 mb-3">
                  Supports standard EIL .xlsx or .xls files (up to 50MB)
                </div>

                <button
                  id={`browse-btn-${type}`}
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-medium text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-md border border-blue-200 transition shadow-2xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  Browse Workbook
                </button>
              </>
            )}

            {/* Expected Fields Checklist */}
            <div className="mt-4 pt-3 border-t border-slate-200 w-full text-left">
              <span className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider block mb-1.5">
                Key Fields Evaluated:
              </span>
              <div className="flex flex-wrap gap-1">
                {info.keyFields.map((h, i) => (
                  <span
                    key={i}
                    className="inline-block px-1.5 py-0.5 text-[10px] bg-slate-100 text-slate-600 rounded border border-slate-200"
                  >
                    {h}
                  </span>
                ))}
              </div>
            </div>
          </div>
        ) : (
          /* Uploaded Workbook State */
          <div className="space-y-3 flex-1 flex flex-col justify-between">
            {/* File Info Box */}
            <div className="bg-slate-50 border border-slate-200 rounded-lg p-3">
              <div className="flex items-start justify-between gap-2">
                <div className="flex items-start gap-2.5 min-w-0">
                  <FileSpreadsheet className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
                  <div className="min-w-0">
                    <div className="text-xs font-bold text-slate-800 truncate" title={metadata.fileName}>
                      {metadata.fileName}
                    </div>
                    <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-2">
                      <span>{formatFileSize(metadata.fileSizeBytes)}</span>
                      <span>•</span>
                      <span>
                        {new Date(metadata.uploadedAt).toLocaleTimeString([], {
                          hour: '2-digit',
                          minute: '2-digit',
                        })}
                      </span>
                      {metadata.isSampleData && (
                        <>
                          <span>•</span>
                          <span className="text-[10px] font-medium bg-blue-100 text-blue-700 px-1 rounded">
                            Sample Data
                          </span>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Metrics Grid */}
              <div className="grid grid-cols-3 gap-2 mt-3 pt-2.5 border-t border-slate-200/80 text-center">
                <div className="bg-white p-1.5 rounded border border-slate-200/60">
                  <div className="text-[10px] text-slate-500 font-medium">Worksheets</div>
                  <div className="text-xs font-bold text-slate-800 mt-0.5">
                    {metadata.sheets.length}
                  </div>
                </div>
                <div className="bg-white p-1.5 rounded border border-slate-200/60">
                  <div className="text-[10px] text-slate-500 font-medium">Data Rows</div>
                  <div className="text-xs font-bold text-slate-800 mt-0.5">
                    {activeSheetData?.rowCount ?? 0}
                  </div>
                </div>
                <div className="bg-white p-1.5 rounded border border-slate-200/60">
                  <div className="text-[10px] text-slate-500 font-medium">Columns</div>
                  <div className="text-xs font-bold text-slate-800 mt-0.5">
                    {activeSheetData?.columnCount ?? 0}
                  </div>
                </div>
              </div>

              {/* Detected Worksheet Name */}
              <div className="mt-2 text-[11px] text-slate-600 flex items-center gap-1 truncate">
                <span className="font-semibold text-slate-500">Active Sheet:</span>
                <span className="font-mono bg-white px-1.5 py-0.5 rounded border border-slate-200 text-slate-700 truncate">
                  {metadata.activeSheet || 'Sheet1'}
                </span>
              </div>
            </div>

            {/* Validation Feedback */}
            {metadata.validationErrors.length > 0 && (
              <div className="p-2.5 rounded-lg bg-rose-50/90 border border-rose-200 text-xs text-rose-800 space-y-1">
                <div className="font-semibold flex items-center gap-1.5">
                  <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                  Structural Issues Detected:
                </div>
                <ul className="list-disc pl-4 space-y-0.5 text-[11px]">
                  {metadata.validationErrors.map((err, i) => (
                    <li key={i}>{err}</li>
                  ))}
                </ul>
              </div>
            )}

            {metadata.validationWarnings.length > 0 && (
              <div className="p-2 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800 space-y-0.5">
                <div className="font-semibold flex items-center gap-1 text-[11px]">
                  <AlertTriangle className="w-3.5 h-3.5 shrink-0 text-amber-600" />
                  Notice for Auditor:
                </div>
                <ul className="list-disc pl-4 text-[11px] text-amber-700">
                  {metadata.validationWarnings.map((w, i) => (
                    <li key={i}>{w}</li>
                  ))}
                </ul>
              </div>
            )}

            {metadata.validationStatus === 'VALID' && (
              <div className="p-2 rounded-lg bg-emerald-50 border border-emerald-200 text-[11px] text-emerald-800 flex items-center gap-1.5 font-medium">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                <span>Workbook structure passed initial validation checks.</span>
              </div>
            )}

            {/* Actions: Inspect, Replace File, Remove (Explicitly Distinguished) */}
            <div className="pt-2 flex items-center gap-2">
              <button
                id={`preview-btn-${type}`}
                type="button"
                onClick={() => onOpenPreview(metadata)}
                className="flex-1 inline-flex items-center justify-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md transition shadow-2xs"
                title="Inspect worksheets, headers, and sample data"
              >
                <Eye className="w-3.5 h-3.5 text-slate-500" />
                <span>Inspect</span>
              </button>

              <button
                id={`replace-btn-${type}`}
                type="button"
                onClick={handleTriggerReplace}
                className="inline-flex items-center justify-center gap-1 px-2.5 py-1.5 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-md transition"
                title="Directly select a new file to replace this workbook"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>Replace File</span>
              </button>

              {isConfirmingDelete ? (
                <div className="flex items-center gap-1 bg-rose-50 border border-rose-200 px-2 py-1 rounded-md text-xs">
                  <span className="text-[11px] text-rose-700 font-medium">Remove?</span>
                  <button
                    type="button"
                    onClick={() => {
                      setIsConfirmingDelete(false);
                      onDelete();
                    }}
                    className="px-2 py-0.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-[10px] font-bold"
                  >
                    Yes
                  </button>
                  <button
                    type="button"
                    onClick={() => setIsConfirmingDelete(false)}
                    className="px-1.5 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-[10px]"
                  >
                    Cancel
                  </button>
                </div>
              ) : (
                <button
                  id={`delete-btn-${type}`}
                  type="button"
                  onClick={() => setIsConfirmingDelete(true)}
                  className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-200 rounded-md transition"
                  title="Remove workbook from active session"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              )}

              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.xlsm"
                className="hidden"
                onChange={handleFileChange}
              />
            </div>
          </div>
        )}

        {/* Upload Error Banner */}
        {uploadError && (
          <div className="mt-3 p-2 bg-rose-50 border border-rose-200 rounded text-xs text-rose-700 flex items-start gap-1.5">
            <AlertCircle className="w-4 h-4 text-rose-500 shrink-0 mt-0.5" />
            <span>{uploadError}</span>
          </div>
        )}
      </div>
    </div>
  );
};
