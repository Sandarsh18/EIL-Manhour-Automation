import React, { useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  Calculator,
  Calendar,
  CheckCircle2,
  AlertTriangle,
  AlertCircle,
  FileText,
  HelpCircle,
  Layers,
  Sparkles,
  RefreshCw,
  Info,
} from 'lucide-react';
import { JobCalculationResult, CalculationLineageItem } from '../types/eil.ts';
import { CalculationEvidenceModal } from './CalculationEvidenceModal';
import { isPresentValue, formatDisplayValue, getValueBadgeClass } from '../utils/eilHighlight';

interface Step4RuleAnalysisViewProps {
  calculations: Record<string, JobCalculationResult>;
  selectedJobs: string[];
  evaluationMonth: string;
  onEvaluationMonthChange: (month: string) => void;
  onRunAnalysis: () => Promise<void>;
  onBackToStep3: () => void;
  onProceedToStep5: () => void;
  isLoading: boolean;
}

export const Step4RuleAnalysisView: React.FC<Step4RuleAnalysisViewProps> = ({
  calculations,
  selectedJobs,
  evaluationMonth,
  onEvaluationMonthChange,
  onRunAnalysis,
  onBackToStep3,
  onProceedToStep5,
  isLoading,
}) => {
  const [activeLineageJob, setActiveLineageJob] = useState<JobCalculationResult | null>(null);
  const [evidenceTab, setEvidenceTab] = useState<'summary' | 'excel1' | 'excel2' | 'lineage'>('summary');

  const calcList = ((Object.values(calculations) || []) as JobCalculationResult[]).filter(
    (c) => c && c.jobNumber && c.jobNumber.trim() !== ''
  );
  const totalAnalyzed = calcList.length;
  const validCount = calcList.filter((c) => c.status !== 'BLOCKED').length;

  // Compute total hours across all calculated jobs
  const totalHoursSum = calcList.reduce((sum, c) => sum + (c.total || 0), 0);

  // Derive OCS Window string for display: 1st through last day of active evaluation month
  const [yearStr, monthStr] = evaluationMonth.split('-');
  const y = parseInt(yearStr, 10);
  const m = parseInt(monthStr, 10);
  const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const windowStartStr = `${yearStr}-${monthStr}-01`;
  const windowEndStr = `${yearStr}-${monthStr}-${String(lastDay).padStart(2, '0')}`;

  return (
    <div className="space-y-6 pb-24">
      {/* Header Bar */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep3}
            className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition"
            title="Return to Step 3 Job Selection"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                Step 4 of 6
              </span>
              <h2 className="text-base font-bold text-slate-900">
                EIL Business Rule Analysis Engine
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Execute deterministic business rules (OCS Date window, valid working days × 8h, QAP logic) with verifiable lineage.
            </p>
          </div>
        </div>

        {/* Evaluation Month Control */}
        <div className="flex items-center gap-3 bg-slate-50 p-2 rounded-lg border border-slate-200">
          <div className="flex items-center gap-1.5 text-xs text-slate-600 font-semibold">
            <Calendar className="w-4 h-4 text-blue-600" />
            <span>Evaluation Period:</span>
          </div>
          <input
            id="input-eval-month"
            type="month"
            value={evaluationMonth}
            onChange={(e) => onEvaluationMonthChange(e.target.value)}
            className="text-xs font-mono bg-white border border-slate-300 rounded px-2.5 py-1 text-slate-800 font-medium"
          />
          <button
            id="btn-re-run-analysis"
            onClick={onRunAnalysis}
            disabled={isLoading}
            className="px-3 py-1 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded transition flex items-center gap-1 shadow-2xs"
          >
            {isLoading ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Calculator className="w-3.5 h-3.5" />}
            <span>Run Analysis</span>
          </button>
        </div>
      </div>

      {/* OCS Window Banner */}
      <div className="bg-blue-50/70 border border-blue-200 rounded-xl p-4 text-xs text-blue-900 flex flex-wrap items-center justify-between gap-3 shadow-2xs">
        <div className="flex items-center gap-2">
          <Sparkles className="w-4 h-4 text-blue-600 shrink-0" />
          <span>
            <strong>Active OCS Eligibility Window:</strong> For evaluation month{' '}
            <span className="font-mono font-bold underline">{evaluationMonth}</span>, OCS completion is awarded
            for records whose OCS Date falls within the month ({windowStartStr} through {windowEndStr}).
          </span>
        </div>
        <span className="text-[11px] text-blue-700 font-medium bg-white px-2 py-0.5 rounded border border-blue-200">
          Selected Month OCS Logic
        </span>
      </div>

      {/* KPIs Summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="bg-white rounded-xl p-3.5 border border-slate-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-slate-500 uppercase">Analyzed Jobs</span>
          <div className="text-xl font-bold text-slate-900 mt-1">{totalAnalyzed}</div>
          <span className="text-[10px] text-slate-400">Total in active scope</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-emerald-200 bg-emerald-50/20 shadow-2xs">
          <span className="text-[11px] font-semibold text-emerald-800 uppercase flex items-center gap-1">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
            Calculated Jobs
          </span>
          <div className="text-xl font-bold text-emerald-800 mt-1">{validCount}</div>
          <span className="text-[10px] text-emerald-600">Calculated successfully</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-indigo-200 bg-indigo-50/20 shadow-2xs">
          <span className="text-[11px] font-semibold text-indigo-800 uppercase flex items-center gap-1">
            <Layers className="w-3.5 h-3.5 text-indigo-600" />
            Selected Scope
          </span>
          <div className="text-xl font-bold text-indigo-900 mt-1">{selectedJobs.length}</div>
          <span className="text-[10px] text-indigo-600">Jobs selected from Excel 3</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-blue-200 bg-blue-50/20 shadow-2xs">
          <span className="text-[11px] font-semibold text-blue-800 uppercase">Aggregated Manhours</span>
          <div className="text-xl font-bold text-blue-900 mt-1 font-mono">{(Number(totalHoursSum) || 0).toFixed(1)}h</div>
          <span className="text-[10px] text-blue-600">Total output hours sum</span>
        </div>
      </div>

      {/* Calculations Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Layers className="w-4 h-4 text-blue-600" />
            <h3 className="text-sm font-bold text-slate-900">
              Rule Execution Results by Job ({calcList.length})
            </h3>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            {/* Visual Guide / Legend */}
            <div className="flex items-center gap-2 text-[11px] text-slate-500 bg-slate-50 px-2.5 py-1 rounded-md border border-slate-200">
              <span className="font-semibold text-slate-700">Audit Highlighting:</span>
              <span className="inline-flex items-center gap-1">
                <span className="inline-block w-2.5 h-2.5 rounded bg-amber-100 border border-amber-300"></span>
                <span className="text-slate-700 font-medium">Actual Data (&gt;0)</span>
              </span>
              <span className="text-slate-300">|</span>
              <span className="inline-flex items-center gap-1">
                <span className="inline-block w-2.5 h-2.5 rounded bg-transparent border border-slate-200"></span>
                <span className="text-slate-400">Zero / Empty</span>
              </span>
            </div>
            <span className="text-xs text-slate-500 hidden sm:inline">
              Click any row or "Lineage" to inspect calculation proof
            </span>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/90 text-slate-700 border-b border-slate-200 font-semibold text-[11px] uppercase tracking-wider">
              <tr>
                <th className="px-3 py-2.5">Job No.</th>
                <th className="px-3 py-2.5">FOI Number</th>
                <th className="px-3 py-2.5 text-center">Origin</th>
                <th className="px-3 py-2.5 text-right">FD</th>
                <th className="px-3 py-2.5 text-right">Running</th>
                <th className="px-3 py-2.5 text-right">OCS Done</th>
                <th className="px-3 py-2.5 text-right">Expediting</th>
                <th className="px-3 py-2.5 text-right">Inspection</th>
                <th className="px-3 py-2.5 text-right">Others</th>
                <th className="px-3 py-2.5 text-right">Meeting</th>
                <th className="px-3 py-2.5 text-right bg-blue-50/50 font-bold text-blue-900">Total (h)</th>
                <th className="px-3 py-2.5 text-center">Status</th>
                <th className="px-3 py-2.5 text-center">Lineage</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
              {calcList.length === 0 ? (
                <tr>
                  <td colSpan={13} className="px-4 py-12 text-center text-slate-400 font-sans">
                    No calculations present. Click "Run Analysis" above to process jobs.
                  </td>
                </tr>
              ) : (
                calcList.map((job, idx) => (
                  <tr
                    key={`calc-${job.jobNumber || 'item'}-${idx}`}
                    id={`calc-row-${job.jobNumber || idx}`}
                    onClick={() => {
                      setEvidenceTab('summary');
                      setActiveLineageJob(job);
                    }}
                    className="hover:bg-slate-50/80 transition cursor-pointer"
                  >
                    <td className="px-3 py-2 font-bold text-slate-900">{job.jobNumber}</td>
                    <td className="px-3 py-2 font-mono text-slate-600 text-xs">
                      {job.foiNumber ? (
                        <span className="font-semibold text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-200">
                          {job.foiNumber}
                        </span>
                      ) : (
                        <span className="text-slate-400 italic">Unavailable</span>
                      )}
                    </td>
                    <td className="px-3 py-2 text-center font-sans">
                      <span
                        className={`inline-flex items-center text-[9px] font-semibold px-2 py-0.5 rounded-full ${
                          job.origin === 'Both'
                            ? 'bg-emerald-100 text-emerald-800'
                            : job.origin === 'PM Call Logs'
                            ? 'bg-purple-100 text-purple-800'
                            : 'bg-blue-100 text-blue-800'
                        }`}
                      >
                        {job.origin || 'Consolidated Report'}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.fd)}>
                        {formatDisplayValue(job.fd)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.runningOrders)}>
                        {formatDisplayValue(job.runningOrders)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.ocsDone)}>
                        {formatDisplayValue(job.ocsDone)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.expediting)}>
                        {formatDisplayValue(job.expediting)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.inspection)}>
                        {formatDisplayValue(job.inspection)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.others)}>
                        {formatDisplayValue(job.others)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right">
                      <span className={getValueBadgeClass(job.meeting)}>
                        {formatDisplayValue(job.meeting)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right bg-blue-50/30">
                      <span className={getValueBadgeClass(job.total, true)}>
                        {formatDisplayValue(job.total)}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-center font-sans" onClick={(e) => e.stopPropagation()}>
                      {job.status === 'BLOCKED' ? (
                        <button
                          onClick={() => {
                            setEvidenceTab('summary');
                            setActiveLineageJob(job);
                          }}
                          className="inline-flex items-center gap-0.5 text-[10px] font-bold text-rose-700 bg-rose-100 hover:bg-rose-200 px-2 py-0.5 rounded-full transition"
                          title="Click to view details"
                        >
                          Blocked
                        </button>
                      ) : (
                        <button
                          onClick={() => {
                            setEvidenceTab('summary');
                            setActiveLineageJob(job);
                          }}
                          className="inline-flex items-center gap-0.5 text-[10px] font-bold text-emerald-700 bg-emerald-100 hover:bg-emerald-200 px-2 py-0.5 rounded-full transition"
                          title="Click to view calculation evidence"
                        >
                          Calculated
                        </button>
                      )}
                    </td>
                    <td className="px-3 py-2 text-center" onClick={(e) => e.stopPropagation()}>
                      <button
                        onClick={() => {
                          setEvidenceTab('lineage');
                          setActiveLineageJob(job);
                        }}
                        className="inline-flex items-center gap-1 px-2 py-1 text-[10px] font-sans font-semibold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 rounded border border-blue-200 transition"
                      >
                        <FileText className="w-3 h-3" />
                        Lineage
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sticky Action Footer */}
      <div className="sticky bottom-0 z-40 bg-white/95 backdrop-blur-md rounded-xl p-4 border border-slate-200 shadow-lg flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <span className="text-xs font-bold text-slate-800">
            {calcList.length} jobs analyzed for {evaluationMonth}
          </span>
          <span className="text-xs text-emerald-600 font-medium">
            ({validCount} jobs calculated)
          </span>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep3}
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition"
          >
            ← Back to Step 3
          </button>

          <button
            id="btn-proceed-step5"
            onClick={onProceedToStep5}
            disabled={calcList.length === 0}
            className={`px-5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-2 shadow-xs ${
              calcList.length > 0
                ? 'bg-blue-600 hover:bg-blue-700 text-white cursor-pointer'
                : 'bg-slate-200 text-slate-400 border border-slate-300 cursor-not-allowed'
            }`}
          >
            <span>Proceed to Step 5: Review & Approve →</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Authoritative Calculation Evidence & Lineage Modal */}
      {activeLineageJob && (
        <CalculationEvidenceModal
          job={activeLineageJob}
          evaluationMonth={evaluationMonth}
          initialTab={evidenceTab}
          onClose={() => setActiveLineageJob(null)}
        />
      )}
    </div>
  );
};
