import React, { useState } from 'react';
import {
  ArrowLeft,
  Download,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  Layers,
  ShieldCheck,
  RefreshCw,
  Sparkles,
  Printer,
  History,
} from 'lucide-react';
import { OutputValidationReport, ReviewItem } from '../types/eil.ts';
import { EilApi } from '../services/api.ts';
import { formatDisplayValue, getValueBadgeClass } from '../utils/eilHighlight';

interface Step6OutputViewProps {
  outputReport: OutputValidationReport | null;
  reviewItems: Record<string, ReviewItem>;
  evaluationMonth: string;
  onGenerateOutput: () => Promise<void>;
  onBackToStep5: () => void;
  isGenerating: boolean;
}

export const Step6OutputView: React.FC<Step6OutputViewProps> = ({
  outputReport,
  reviewItems,
  evaluationMonth,
  onGenerateOutput,
  onBackToStep5,
  isGenerating,
}) => {
  const items = ((Object.values(reviewItems) || []) as ReviewItem[]).filter(
    (i) => i && i.jobNumber && i.jobNumber.trim() !== ''
  );
  const approvedItems = items.filter((i) => i.approvalStatus === 'APPROVED');
  const overriddenItems = items.filter((i) => i.isOverridden);
  const totalHours = approvedItems.reduce((sum, i) => sum + (i.total || 0), 0);

  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);

  const handleDownloadExcel = async () => {
    try {
      setIsDownloading(true);
      setDownloadError(null);
      const { blob, fileName } = await EilApi.downloadOutputBlob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName || `EIL_Manhour_Audit_Output_${evaluationMonth}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('Download failed:', err);
      setDownloadError(err.message || 'Output workbook generation failed. No output file is available.');
    } finally {
      setIsDownloading(false);
    }
  };

  const handleExportAuditJson = () => {
    const auditData = {
      evaluationMonth,
      generatedAt: new Date().toISOString(),
      summary: {
        totalJobs: items.length,
        approvedJobs: approvedItems.length,
        overriddenJobs: overriddenItems.length,
        totalManhours: totalHours,
      },
      validationChecklist: outputReport?.checks,
      records: items.map((item) => ({
        jobNumber: item.jobNumber,
        manhours: {
          fd: item.fd,
          runningOrders: item.runningOrders,
          ocsDone: item.ocsDone,
          expediting: item.expediting,
          inspection: item.inspection,
          others: item.others,
          meeting: item.meeting,
          total: item.total,
        },
        status: item.approvalStatus,
        isOverridden: item.isOverridden,
        auditNotes: item.auditNotes,
        lineage: item.lineage,
      })),
    };

    const blob = new Blob([JSON.stringify(auditData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `EIL_Audit_Trail_${evaluationMonth}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-6 pb-24">
      {downloadError && (
        <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-800 text-xs flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
            <span>{downloadError}</span>
          </div>
          <button
            onClick={() => setDownloadError(null)}
            className="text-xs font-semibold text-rose-600 hover:text-rose-800 underline ml-4"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* Header Bar */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep5}
            className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition"
            title="Return to Step 5 Review"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                Step 6 of 6
              </span>
              <h2 className="text-base font-bold text-slate-900">
                Master Workbook Output & Validation
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Outputs updated calculations for jobs related to Excel 1 and Excel 2 into the Master Template while ignoring pre-existing uploaded data.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            id="btn-regenerate-output"
            onClick={onGenerateOutput}
            disabled={isGenerating}
            className="px-4 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition flex items-center gap-1.5 shadow-2xs"
          >
            {isGenerating ? (
              <RefreshCw className="w-3.5 h-3.5 animate-spin" />
            ) : (
              <Layers className="w-3.5 h-3.5" />
            )}
            <span>Regenerate Output</span>
          </button>
        </div>
      </div>

      {/* Primary Output Status Card */}
      <div className="glass-card rounded-2xl p-6 sm:p-7 space-y-6">
        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200/80 pb-6">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-600 text-white flex items-center justify-center shadow-md shadow-emerald-500/20 ring-4 ring-emerald-100/70">
              <FileSpreadsheet className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h3 className="text-lg font-bold text-slate-900">
                  {outputReport?.reportMetadata?.fileName || `EIL_Master_Audit_Output_${evaluationMonth}.xlsx`}
                </h3>
                <span className="inline-flex items-center gap-1 text-[11px] font-bold text-emerald-800 bg-emerald-100/80 border border-emerald-300 px-2.5 py-0.5 rounded-full shadow-2xs">
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                  Ready for Download
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-1">
                Generated with verified job calculations populated directly into the Excel 3 master template.
              </p>
              {outputReport?.reportMetadata?.generatedAt && (
                <div className="flex items-center gap-3 text-[11px] text-slate-400 mt-1.5 font-mono">
                  <span>Generated: {new Date(outputReport.reportMetadata.generatedAt).toLocaleString()}</span>
                  {outputReport?.outputFile?.fileSizeBytes != null && !isNaN(Number(outputReport.outputFile.fileSizeBytes)) && (
                    <span>• Size: {(Number(outputReport.outputFile.fileSizeBytes) / 1024).toFixed(1)} KB</span>
                  )}
                </div>
              )}
            </div>
          </div>

          <div className="flex items-center gap-2.5 flex-wrap">
            <button
              id="btn-download-master-excel"
              onClick={handleDownloadExcel}
              disabled={!outputReport || isDownloading}
              className="px-5 py-2.5 text-xs font-bold text-white bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl transition flex items-center gap-2 shadow-md shadow-emerald-600/20 cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>{isDownloading ? 'Downloading...' : 'Download Master Workbook (.xlsx)'}</span>
            </button>

            <button
              onClick={handleExportAuditJson}
              className="px-4 py-2.5 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-xl transition flex items-center gap-2 shadow-2xs cursor-pointer"
              title="Download structured audit report"
            >
              <History className="w-4 h-4 text-slate-500" />
              <span>Export Audit Trail (JSON)</span>
            </button>
          </div>
        </div>

        {/* Change & Impact Summary Metrics */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-500">Populated Jobs</span>
            <div className="text-xl font-bold text-slate-900 mt-1">
              {outputReport?.totalJobsUpdated ?? approvedItems.length}
            </div>
            <span className="text-[10px] text-slate-400">Written to template rows</span>
          </div>

          <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-500">Approved Jobs</span>
            <div className="text-xl font-bold text-emerald-700 mt-1">
              {approvedItems.length}
            </div>
            <span className="text-[10px] text-emerald-600">Verified in Step 5</span>
          </div>

          <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-500">Manual Overrides</span>
            <div className="text-xl font-bold text-indigo-700 mt-1">{overriddenItems.length}</div>
            <span className="text-[10px] text-indigo-600">Documented with notes</span>
          </div>

          <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-500">Total Manhours Output</span>
            <div className="text-xl font-bold text-blue-700 mt-1 font-mono">{(Number(totalHours) || 0).toFixed(1)}h</div>
            <span className="text-[10px] text-blue-600">Calculated sum</span>
          </div>
        </div>
      </div>

      {/* Final Verification Table (Summary Preview) */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
          <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">
            Output Preview — Top Verified Records
          </h3>
          <span className="text-xs text-slate-500">
            {approvedItems.length} jobs written to Master Template
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 font-semibold uppercase text-[11px]">
              <tr>
                <th className="px-3 py-2.5">Job Number</th>
                <th className="px-3 py-2.5">FOI Number</th>
                <th className="px-3 py-2.5 text-center">Origin</th>
                <th className="px-3 py-2.5 text-right">FD</th>
                <th className="px-3 py-2.5 text-right">Running</th>
                <th className="px-3 py-2.5 text-right">OCS Done</th>
                <th className="px-3 py-2.5 text-right">Expediting</th>
                <th className="px-3 py-2.5 text-right">Inspection</th>
                <th className="px-3 py-2.5 text-right">Others</th>
                <th className="px-3 py-2.5 text-right">Meeting</th>
                <th className="px-3 py-2.5 text-right font-bold text-blue-900 bg-blue-50/50">Total</th>
                <th className="px-3 py-2.5 text-center">Audit Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
              {approvedItems.slice(0, 15).map((item, idx) => (
                <tr key={`out-job-${item.jobNumber || 'job'}-${idx}`} className="hover:bg-slate-50/80">
                  <td className="px-3 py-2 font-bold text-slate-900">{item.jobNumber}</td>
                  <td className="px-3 py-2 font-mono text-slate-600 text-xs">
                    {item.foiNumber ? (
                      <span className="font-semibold text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-200">
                        {item.foiNumber}
                      </span>
                    ) : (
                      <span className="text-slate-400 italic">Unavailable</span>
                    )}
                  </td>
                  <td className="px-3 py-2 text-center font-sans">
                    <span
                      className={`inline-flex items-center text-[9px] font-semibold px-2 py-0.5 rounded-full ${
                        item.origin === 'Both'
                          ? 'bg-emerald-100 text-emerald-800'
                          : item.origin === 'PM Call Logs'
                          ? 'bg-purple-100 text-purple-800'
                          : 'bg-blue-100 text-blue-800'
                      }`}
                    >
                      {item.origin || 'Consolidated Report'}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.fd)}>{formatDisplayValue(item.fd)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.runningOrders)}>{formatDisplayValue(item.runningOrders)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.ocsDone)}>{formatDisplayValue(item.ocsDone)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.expediting)}>{formatDisplayValue(item.expediting)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.inspection)}>{formatDisplayValue(item.inspection)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.others)}>{formatDisplayValue(item.others)}</span>
                  </td>
                  <td className="px-3 py-2 text-right">
                    <span className={getValueBadgeClass(item.meeting)}>{formatDisplayValue(item.meeting)}</span>
                  </td>
                  <td className="px-3 py-2 text-right bg-blue-50/30">
                    <span className={getValueBadgeClass(item.total, true)}>{formatDisplayValue(item.total)}</span>
                  </td>
                  <td className="px-3 py-2 text-center font-sans">
                    <span className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded-full">
                      <CheckCircle2 className="w-3 h-3" />
                      Audited
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sticky Action Footer */}
      <div className="sticky bottom-0 z-40 bg-white/95 backdrop-blur-md rounded-xl p-4 border border-slate-200 shadow-lg flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <span className="text-xs font-bold text-slate-800">
            Workflow Complete: Master output workbook ready for download
          </span>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep5}
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition"
          >
            ← Back to Step 5 Review
          </button>

          <button
            onClick={handleDownloadExcel}
            disabled={!outputReport || isDownloading}
            className="px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg transition flex items-center gap-2 shadow-xs cursor-pointer"
          >
            <Download className="w-4 h-4" />
            <span>{isDownloading ? 'Downloading...' : 'Download Output Workbook'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};
