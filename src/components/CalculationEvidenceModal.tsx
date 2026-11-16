import React, { useState } from 'react';
import {
  X,
  FileText,
  CheckCircle2,
  AlertTriangle,
  AlertOctagon,
  Layers,
  Calculator,
  FileSpreadsheet,
  Calendar,
  Info,
  ChevronRight,
  Filter,
  ArrowRight,
  ShieldCheck,
  Building2,
  Check,
  Table,
} from 'lucide-react';
import { JobCalculationResult, ReviewItem, CalculationLineageItem, Excel1OrderRecord, Excel2InspectionRecord, WarningAuditDetail } from '../types/eil';
import { isPresentValue, formatDisplayValue, getValueBadgeClass } from '../utils/eilHighlight';

export interface CalculationEvidenceModalProps {
  job: JobCalculationResult | ReviewItem | null;
  evaluationMonth?: string;
  initialTab?: 'summary' | 'excel1' | 'excel2' | 'lineage';
  onEditOverride?: (job: JobCalculationResult | ReviewItem) => void;
  onClose: () => void;
}

export const CalculationEvidenceModal: React.FC<CalculationEvidenceModalProps> = ({
  job,
  evaluationMonth = 'Current Cycle',
  initialTab = 'summary',
  onEditOverride,
  onClose,
}) => {
  const [activeTab, setActiveTab] = useState<'summary' | 'excel1' | 'excel2' | 'lineage'>(initialTab);
  const [excel1Filter, setExcel1Filter] = useState<'all' | 'running' | 'fd' | 'ocs' | 'excluded'>('all');
  const [excel2Filter, setExcel2Filter] = useState<'all' | 'inspection' | 'qap' | 'review' | 'conflict' | 'excluded'>('all');

  if (!job) return null;

  // Safe extraction with default fallbacks
  const jobNumber = job.jobNumber || 'Unknown Job';
  const expediting = Number(job.expediting ?? 0);
  const inspection = Number(job.inspection ?? 0);
  const others = Number(job.others ?? 0);
  const meeting = Number(job.meeting ?? 0);
  const total = Number(job.total ?? 0);
  const fd = Number(job.fd ?? 0);
  const runningOrders = Number(job.runningOrders ?? 0);
  const ocsDone = Number(job.ocsDone ?? 0);

  const lineage: CalculationLineageItem[] = job.lineage || [];
  const sourceOrders: Excel1OrderRecord[] = job.sourceOrdersExcel1 || [];
  const sourceCalls: Excel2InspectionRecord[] = job.sourceCallsExcel2 || [];

  const isApproved = 'approvalStatus' in job && job.approvalStatus === 'APPROVED';
  const isOverridden = 'isOverridden' in job && job.isOverridden;

  // Excel 1 filter logic
  const filteredOrders = sourceOrders.filter((ord) => {
    if (excel1Filter === 'all') return true;
    if (excel1Filter === 'running') return ord.classification === 'RUNNING_ORDER';
    if (excel1Filter === 'fd') return ord.classification === 'ORDERS_FD';
    if (excel1Filter === 'ocs') return ord.classification === 'OCS_DONE' || ord.classification === 'EXCLUDED_FD';
    if (excel1Filter === 'excluded') return ord.eligibility === 'Excluded' || ord.classification === 'INVALID';
    return true;
  });

  // Excel 2 filter logic
  const filteredCalls = sourceCalls.filter((call) => {
    if (excel2Filter === 'all') return true;
    if (excel2Filter === 'inspection') return call.status === 'INCLUDED_INSPECTION' || call.status === 'INCLUDED_INSPECTION_WITH_CONFLICT';
    if (excel2Filter === 'qap') return call.status === 'INCLUDED_QAP' || call.status === 'INCLUDED_QAP_WITH_CONFLICT';
    if (excel2Filter === 'review') return call.status === 'INCLUDED_REVIEW';
    if (excel2Filter === 'conflict') return call.status.includes('CONFLICT') || call.status === 'AMBIGUOUS';
    if (excel2Filter === 'excluded') return call.eligibility === 'Excluded' || call.status === 'EXCLUDED_DATES' || call.status === 'ZERO_DAYS' || call.status === 'NO_ACTIVITY' || call.status === 'NO_DATES';
    return true;
  });

  // Counts for tabs/badges
  const excel1EligibleCount = sourceOrders.filter((o) => o.eligibility === 'Eligible').length;
  const excel1ExcludedCount = sourceOrders.length - excel1EligibleCount;
  const excel2EligibleCount = sourceCalls.filter((c) => c.eligibility === 'Eligible').length;
  const excel2ExcludedCount = sourceCalls.length - excel2EligibleCount;

  return (
    <div
      id="calculation-evidence-modal-backdrop"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 backdrop-blur-xs p-3 sm:p-5 overflow-y-auto"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      role="dialog"
      aria-modal="true"
      aria-labelledby="evidence-modal-title"
    >
      <div
        id="calculation-evidence-modal-container"
        className="relative w-full max-w-5xl max-h-[92vh] flex flex-col bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden text-slate-800 animate-in fade-in zoom-in-95 duration-150"
      >
        {/* Header */}
        <div
          id="evidence-modal-header"
          className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50/90 shrink-0"
        >
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-100/80 rounded-lg text-blue-900">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 id="evidence-modal-title" className="text-lg font-bold text-slate-900 tracking-tight">
                  Calculation Evidence & Lineage — Job {jobNumber}
                </h2>
                {isOverridden ? (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold bg-indigo-100 text-indigo-800 border border-indigo-200">
                    Overridden
                  </span>
                ) : isApproved ? (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
                    <CheckCircle2 className="w-3 h-3" /> Approved
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold bg-blue-100 text-blue-800 border border-blue-200">
                    <CheckCircle2 className="w-3 h-3" /> Calculated
                  </span>
                )}
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Authoritative source-by-source trace for evaluation cycle: <span className="font-semibold text-slate-700">{evaluationMonth}</span>
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {onEditOverride && (
              <button
                id="modal-edit-override-btn"
                onClick={() => onEditOverride(job)}
                className="px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition flex items-center gap-1.5"
              >
                <span>Edit Override</span>
              </button>
            )}
            <button
              id="close-evidence-modal-btn"
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors"
              title="Close Audit Evidence (Esc)"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Top Metric Strip / Summary Cards */}
        <div id="evidence-modal-summary-strip" className="grid grid-cols-2 sm:grid-cols-5 gap-2 px-6 py-3.5 bg-slate-100/60 border-b border-slate-200 shrink-0">
          {/* Expediting Card */}
          <div className="p-2.5 rounded-lg bg-white border border-slate-200 shadow-2xs">
            <div className="text-[11px] font-medium text-slate-500 uppercase tracking-wider">Expediting</div>
            <div className="mt-1 flex items-baseline justify-between">
              <span className={getValueBadgeClass(expediting)}>
                {formatDisplayValue(expediting)} MH
              </span>
              <span className="text-[11px] text-slate-400 font-mono">2×(RO+OCS)</span>
            </div>
            <div className="text-[10px] text-slate-500 mt-1">
              RO: {runningOrders} | OCS: {ocsDone} | FD: {fd}
            </div>
          </div>

          {/* Inspection Card */}
          <div className="p-2.5 rounded-lg bg-white border border-slate-200 shadow-2xs">
            <div className="text-[11px] font-medium text-slate-500 uppercase tracking-wider">Inspection</div>
            <div className="mt-1 flex items-baseline justify-between">
              <span className={getValueBadgeClass(inspection)}>
                {formatDisplayValue(inspection)} MH
              </span>
              <span className="text-[11px] text-slate-400 font-mono">Days × 8</span>
            </div>
            <div className="text-[10px] text-slate-500 mt-1">
              {sourceCalls.filter((c) => c.status === 'INCLUDED_INSPECTION').length} active call(s)
            </div>
          </div>

          {/* Others Card */}
          <div className="p-2.5 rounded-lg bg-white border border-slate-200 shadow-2xs">
            <div className="text-[11px] font-medium text-slate-500 uppercase tracking-wider">Others (QAP / Review)</div>
            <div className="mt-1 flex items-baseline justify-between">
              <span className={getValueBadgeClass(others)}>
                {formatDisplayValue(others)} MH
              </span>
              <span className="text-[11px] text-slate-400 font-mono">Days × 1</span>
            </div>
            <div className="text-[10px] text-slate-500 mt-1">
              {sourceCalls.filter((c) => c.status === 'INCLUDED_QAP' || c.status === 'INCLUDED_QAP_WITH_CONFLICT').length} QAP | {sourceCalls.filter((c) => c.status === 'INCLUDED_REVIEW').length} Review
            </div>
          </div>

          {/* Meeting Card */}
          <div className="p-2.5 rounded-lg bg-white border border-slate-200 shadow-2xs">
            <div className="text-[11px] font-medium text-slate-500 uppercase tracking-wider">Meeting</div>
            <div className="mt-1 flex items-baseline justify-between">
              <span className={getValueBadgeClass(meeting)}>
                {formatDisplayValue(meeting)} MH
              </span>
              <span className="text-[11px] text-slate-400 font-mono">Hours</span>
            </div>
            <div className="text-[10px] text-slate-500 mt-1">Direct hours logged</div>
          </div>

          {/* Total Card */}
          <div className="p-2.5 rounded-lg bg-blue-50/70 border border-blue-200 shadow-2xs col-span-2 sm:col-span-1">
            <div className="text-[11px] font-bold text-blue-900 uppercase tracking-wider">Total Manhours</div>
            <div className="mt-1 flex items-baseline justify-between">
              <span className={getValueBadgeClass(total, true)}>
                {formatDisplayValue(total)} MH
              </span>
              <span className="text-[11px] text-blue-700 font-bold font-mono">Sum</span>
            </div>
            <div className="text-[10px] text-blue-700 font-medium mt-1">Exp + Insp + Oth + Mtg</div>
          </div>
        </div>

        {/* Navigation Tabs */}
        <div id="evidence-modal-nav-tabs" className="flex items-center gap-1 px-6 pt-2 border-b border-slate-200 bg-white shrink-0 overflow-x-auto">
          <button
            id="tab-btn-summary"
            onClick={() => setActiveTab('summary')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg border-b-2 transition-colors flex items-center gap-1.5 whitespace-nowrap ${
              activeTab === 'summary'
                ? 'border-blue-600 text-blue-700 bg-blue-50/40'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <Calculator className="w-3.5 h-3.5" /> Calculation Formulas & Summary
          </button>

          <button
            id="tab-btn-excel1"
            onClick={() => setActiveTab('excel1')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg border-b-2 transition-colors flex items-center gap-1.5 whitespace-nowrap ${
              activeTab === 'excel1'
                ? 'border-blue-600 text-blue-700 bg-blue-50/40'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <Layers className="w-3.5 h-3.5" /> Excel 1: Orders ({sourceOrders.length})
          </button>

          <button
            id="tab-btn-excel2"
            onClick={() => setActiveTab('excel2')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg border-b-2 transition-colors flex items-center gap-1.5 whitespace-nowrap ${
              activeTab === 'excel2'
                ? 'border-blue-600 text-blue-700 bg-blue-50/40'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" /> Excel 2: Inspection Calls ({sourceCalls.length})
          </button>

          <button
            id="tab-btn-lineage"
            onClick={() => setActiveTab('lineage')}
            className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg border-b-2 transition-colors flex items-center gap-1.5 whitespace-nowrap ${
              activeTab === 'lineage'
                ? 'border-blue-600 text-blue-700 bg-blue-50/40'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <Table className="w-3.5 h-3.5" /> Cell Trace & Rules ({lineage.length})
          </button>
        </div>

        {/* Modal Scrollable Body */}
        <div id="evidence-modal-content-area" className="flex-1 overflow-y-auto p-6 space-y-6 bg-slate-50/50 min-h-[360px]">
          {/* TAB 1: SUMMARY & FORMULA BREAKDOWN */}
          {activeTab === 'summary' && (
            <div className="space-y-6">
              {/* Formula Cards */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Expediting Rule Card */}
                <div className="p-4 rounded-xl bg-white border border-slate-200 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between">
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      <Layers className="w-4 h-4 text-blue-600" /> Expediting Manhours
                    </h3>
                    <span className={getValueBadgeClass(expediting)}>
                      {formatDisplayValue(expediting)} MH
                    </span>
                  </div>
                  <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 font-mono text-xs text-slate-800">
                    Expediting = 2 × (Running Orders + OCS Done)
                  </div>
                  <div className="text-xs text-slate-600 space-y-1">
                    <p>
                      • <strong className="text-slate-800">Running Orders:</strong> {runningOrders} order(s) where Balance Qty ≠ 0
                    </p>
                    <p>
                      • <strong className="text-slate-800">OCS Done:</strong> {ocsDone} order(s) where Balance Qty = 0
                    </p>
                    <p>
                      • <strong className="text-slate-800">Calculation:</strong> 2 × ({runningOrders} + {ocsDone}) = {expediting} MH
                    </p>
                    <p className="text-[11px] text-slate-500 pt-1 border-t border-slate-100">
                      Note: Orders for FD ({fd}) are items where Balance Qty = 0, OCS Date is empty, and Head RPO is blank, tracked for final documentation follow-up without adding extra hours.
                    </p>
                  </div>
                </div>

                {/* Inspection Rule Card */}
                <div className="p-4 rounded-xl bg-white border border-slate-200 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between">
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600" /> Inspection Manhours
                    </h3>
                    <span className={getValueBadgeClass(inspection)}>
                      {formatDisplayValue(inspection)} MH
                    </span>
                  </div>
                  <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 font-mono text-xs text-slate-800">
                    Inspection = Attended Days in {evaluationMonth} × 8
                  </div>
                  <div className="text-xs text-slate-600 space-y-1">
                    <p>
                      • Multiplier of <strong className="text-slate-800">8 hours/day</strong> applied strictly to eligible inspection call records.
                    </p>
                    <p>
                      • Calls with attendance outside {evaluationMonth} or 0 working days are excluded.
                    </p>
                    <p>
                      • <strong className="text-slate-800">Attended calls:</strong> {sourceCalls.filter((c) => c.status === 'INCLUDED_INSPECTION').length} qualified record(s).
                    </p>
                  </div>
                </div>

                {/* Others / QAP & Review Dates Rule Card */}
                <div className="p-4 rounded-xl bg-white border border-slate-200 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between">
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      <ShieldCheck className="w-4 h-4 text-purple-600" /> Others (QAP & Review Dates)
                    </h3>
                    <span className={getValueBadgeClass(others)}>
                      {formatDisplayValue(others)} MH
                    </span>
                  </div>
                  <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 font-mono text-xs text-slate-800">
                    Others = Number of Days × 1 (QAP Priority, then Review Dates)
                  </div>
                  <div className="text-xs text-slate-600 space-y-1">
                    <p>
                      • <strong className="text-slate-800">Case B (QAP):</strong> Triggered when QAP approval is present without physical inspection dates.
                    </p>
                    <p>
                      • <strong className="text-slate-800">Case C (Review Dates):</strong> Triggered when Review dates are present in {evaluationMonth} without physical inspection dates or QAP.
                    </p>
                    <p>
                      • Decimals (e.g. 0.5, 1.5) are preserved exactly and <strong className="text-slate-800">never multiplied by 8</strong>. Overlapping indicators are de-duplicated.
                    </p>
                    <div className="pt-1 flex flex-wrap gap-2 text-[11px] font-medium">
                      <span className="px-2 py-0.5 rounded bg-purple-50 text-purple-800 border border-purple-200">
                        QAP: {sourceCalls.filter((c) => c.status === 'INCLUDED_QAP' || c.status === 'INCLUDED_QAP_WITH_CONFLICT').length}
                      </span>
                      <span className="px-2 py-0.5 rounded bg-indigo-50 text-indigo-800 border border-indigo-200">
                        Review: {sourceCalls.filter((c) => c.status === 'INCLUDED_REVIEW').length}
                      </span>
                      {sourceCalls.some((c) => c.status.includes('CONFLICT')) && (
                        <span className="px-2 py-0.5 rounded bg-amber-50 text-amber-800 border border-amber-200">
                          Resolved Conflicts: {sourceCalls.filter((c) => c.status.includes('CONFLICT')).length}
                        </span>
                      )}
                    </div>
                  </div>
                </div>

                {/* Meeting & Total Card */}
                <div className="p-4 rounded-xl bg-white border border-slate-200 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between">
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      <Calculator className="w-4 h-4 text-blue-700" /> Total Manhour Equation
                    </h3>
                    <span className={getValueBadgeClass(total, true)}>
                      {formatDisplayValue(total)} MH
                    </span>
                  </div>
                  <div className="p-2.5 rounded-lg bg-blue-50/80 border border-blue-200 font-mono text-xs text-blue-950 font-bold">
                    Total = Expediting ({formatDisplayValue(expediting)}) + Inspection ({formatDisplayValue(inspection)}) + Others ({formatDisplayValue(others)}) + Meeting ({formatDisplayValue(meeting)})
                  </div>
                  <div className="text-xs text-slate-600 space-y-1">
                    <p>
                      • <strong className="text-slate-800">Meeting:</strong> {formatDisplayValue(meeting)} MH (carried forward / manually verified).
                    </p>
                    <p>
                      • <strong className="text-slate-800">Authoritative Total:</strong> {formatDisplayValue(total)} Manhours.
                    </p>
                  </div>
                </div>
              </div>

              {/* Source Files Overview */}
              <div className="p-4 rounded-xl bg-white border border-slate-200 shadow-2xs space-y-3">
                <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider">Source Workbooks Loaded</h4>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div className="p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs space-y-1">
                    <div className="font-semibold text-slate-900 flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-blue-500"></span> Excel 1 (Orders)
                    </div>
                    <div className="text-slate-600">Total Rows for Job: <span className="font-semibold">{sourceOrders.length}</span></div>
                    <div className="text-slate-500 text-[11px]">Feeds: FD, Running Orders, OCS Done, Expediting</div>
                  </div>

                  <div className="p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs space-y-1">
                    <div className="font-semibold text-slate-900 flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-emerald-500"></span> Excel 2 (Calls)
                    </div>
                    <div className="text-slate-600">Total Calls for Job: <span className="font-semibold">{sourceCalls.length}</span></div>
                    <div className="text-slate-500 text-[11px]">Feeds: Inspection (×8), Others (QAP ×1)</div>
                  </div>

                  <div className="p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs space-y-1">
                    <div className="font-semibold text-slate-900 flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-indigo-500"></span> Excel 3 (Template)
                    </div>
                    <div className="text-slate-600">Baseline Audit Target: <span className="font-semibold">{evaluationMonth}</span></div>
                    <div className="text-slate-500 text-[11px]">Cross-checks calculated results vs. billing sheet</div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: EXCEL 1 ORDERS EVIDENCE */}
          {activeTab === 'excel1' && (
            <div className="space-y-4">
              {/* Filter and stats row */}
              <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-white rounded-xl border border-slate-200 shadow-2xs">
                <div className="flex items-center gap-2 text-xs">
                  <span className="font-semibold text-slate-700">Filter Orders:</span>
                  <button
                    onClick={() => setExcel1Filter('all')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel1Filter === 'all' ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                  >
                    All ({sourceOrders.length})
                  </button>
                  <button
                    onClick={() => setExcel1Filter('running')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel1Filter === 'running' ? 'bg-amber-600 text-white' : 'bg-amber-50 text-amber-900 border border-amber-200 hover:bg-amber-100'
                    }`}
                  >
                    Running Orders ({runningOrders})
                  </button>
                  <button
                    onClick={() => setExcel1Filter('ocs')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel1Filter === 'ocs' ? 'bg-blue-600 text-white' : 'bg-blue-50 text-blue-900 border border-blue-200 hover:bg-blue-100'
                    }`}
                  >
                    OCS Done ({ocsDone})
                  </button>
                  <button
                    onClick={() => setExcel1Filter('fd')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel1Filter === 'fd' ? 'bg-emerald-600 text-white' : 'bg-emerald-50 text-emerald-900 border border-emerald-200 hover:bg-emerald-100'
                    }`}
                  >
                    Orders for FD ({fd})
                  </button>
                  {excel1ExcludedCount > 0 && (
                    <button
                      onClick={() => setExcel1Filter('excluded')}
                      className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                        excel1Filter === 'excluded' ? 'bg-rose-600 text-white' : 'bg-rose-50 text-rose-800 border border-rose-200 hover:bg-rose-100'
                      }`}
                    >
                      Excluded / Invalid ({excel1ExcludedCount})
                    </button>
                  )}
                </div>

                <div className="text-xs text-slate-500 font-mono">
                  Eligible: <span className="font-bold text-slate-800">{excel1EligibleCount}</span> | Excluded: <span className="font-bold text-slate-800">{excel1ExcludedCount}</span>
                </div>
              </div>

              {/* Table of Orders */}
              {filteredOrders.length === 0 ? (
                <div className="p-8 text-center bg-white rounded-xl border border-slate-200 text-slate-500 text-xs">
                  No order records match the selected filter.
                </div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100/90 text-slate-700 font-semibold border-b border-slate-200">
                          <th className="py-2.5 px-3">#</th>
                          <th className="py-2.5 px-3">FOI Number</th>
                          <th className="py-2.5 px-3">PO Number</th>
                          <th className="py-2.5 px-3">Description / Item</th>
                          <th className="py-2.5 px-3 text-right">Balance Qty</th>
                          <th className="py-2.5 px-3">OCS Date</th>
                          <th className="py-2.5 px-3">Head RPO</th>
                          <th className="py-2.5 px-3">Classification</th>
                          <th className="py-2.5 px-3">Eligibility</th>
                          <th className="py-2.5 px-3">Source Cell</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredOrders.map((ord, idx) => {
                          const isPopBal = isPresentValue(ord.balanceQty);
                          return (
                            <tr key={`ord-${ord.rowIndex || idx}`} className="hover:bg-slate-50/80 transition-colors">
                              <td className="py-2.5 px-3 font-mono text-slate-400 text-[11px]">{idx + 1}</td>
                              <td className="py-2.5 px-3 font-mono font-medium text-slate-900">
                                {ord.foiNumber ? (
                                  <span className="font-semibold text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-200">
                                    {ord.foiNumber}
                                  </span>
                                ) : (
                                  <span className="text-slate-400 italic text-[11px]">Unavailable</span>
                                )}
                              </td>
                              <td className="py-2.5 px-3 font-mono text-slate-700 text-xs">
                                {ord.poNumber || ord.orderNo || '—'}
                              </td>
                              <td className="py-2.5 px-3 text-slate-700 max-w-[200px] truncate" title={ord.description}>
                                {ord.description || '—'}
                              </td>
                              <td className="py-2.5 px-3 text-right">
                                <span className={getValueBadgeClass(ord.balanceQty)}>
                                  {formatDisplayValue(ord.balanceQty, '0')}
                                </span>
                              </td>
                              <td className="py-2.5 px-3 font-mono text-slate-600 text-[11px]">{ord.ocsDate || '—'}</td>
                              <td className="py-2.5 px-3 font-mono text-slate-600 text-[11px]">
                                {ord.headRpoDate ? ord.headRpoDate : <span className="text-slate-300 italic">Blank</span>}
                              </td>
                              <td className="py-2.5 px-3">
                                <span
                                  className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${
                                    ord.classification === 'RUNNING_ORDER'
                                      ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                      : ord.classification === 'ORDERS_FD'
                                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                      : ord.classification === 'OCS_DONE'
                                      ? 'bg-blue-100 text-blue-900 border border-blue-200'
                                      : 'bg-slate-100 text-slate-700 border border-slate-200'
                                  }`}
                                >
                                  {ord.classificationLabel}
                                </span>
                              </td>
                              <td className="py-2.5 px-3">
                                {ord.eligibility === 'Eligible' ? (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700">
                                    <Check className="w-3 h-3" /> Eligible
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700">
                                    <X className="w-3 h-3" /> Excluded
                                  </span>
                                )}
                              </td>
                              <td className="py-2.5 px-3 font-mono text-[11px] text-blue-700 bg-blue-50/40">
                                {ord.sourceCell || `Row ${ord.rowIndex}`}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB 3: EXCEL 2 INSPECTION CALLS EVIDENCE */}
          {activeTab === 'excel2' && (
            <div className="space-y-4">
              {/* Filter and stats row */}
              <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-white rounded-xl border border-slate-200 shadow-2xs">
                <div className="flex flex-wrap items-center gap-2 text-xs">
                  <span className="font-semibold text-slate-700">Filter Calls:</span>
                  <button
                    onClick={() => setExcel2Filter('all')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel2Filter === 'all' ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                  >
                    All ({sourceCalls.length})
                  </button>
                  <button
                    onClick={() => setExcel2Filter('inspection')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel2Filter === 'inspection' ? 'bg-emerald-600 text-white' : 'bg-emerald-50 text-emerald-900 border border-emerald-200 hover:bg-emerald-100'
                    }`}
                  >
                    Inspection (×8) ({sourceCalls.filter((c) => c.status === 'INCLUDED_INSPECTION' || c.status === 'INCLUDED_INSPECTION_WITH_CONFLICT').length})
                  </button>
                  <button
                    onClick={() => setExcel2Filter('qap')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel2Filter === 'qap' ? 'bg-purple-600 text-white' : 'bg-purple-50 text-purple-900 border border-purple-200 hover:bg-purple-100'
                    }`}
                  >
                    Others QAP (×1) ({sourceCalls.filter((c) => c.status === 'INCLUDED_QAP' || c.status === 'INCLUDED_QAP_WITH_CONFLICT').length})
                  </button>
                  <button
                    onClick={() => setExcel2Filter('review')}
                    className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                      excel2Filter === 'review' ? 'bg-indigo-600 text-white' : 'bg-indigo-50 text-indigo-900 border border-indigo-200 hover:bg-indigo-100'
                    }`}
                  >
                    Review Dates (×1) ({sourceCalls.filter((c) => c.status === 'INCLUDED_REVIEW').length})
                  </button>
                  {sourceCalls.filter((c) => c.status.includes('CONFLICT') || c.status === 'AMBIGUOUS').length > 0 && (
                    <button
                      onClick={() => setExcel2Filter('conflict')}
                      className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                        excel2Filter === 'conflict' ? 'bg-amber-600 text-white' : 'bg-amber-50 text-amber-900 border border-amber-200 hover:bg-amber-100'
                      }`}
                    >
                      Conflicts ({sourceCalls.filter((c) => c.status.includes('CONFLICT') || c.status === 'AMBIGUOUS').length})
                    </button>
                  )}
                  {excel2ExcludedCount > 0 && (
                    <button
                      onClick={() => setExcel2Filter('excluded')}
                      className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                        excel2Filter === 'excluded' ? 'bg-rose-600 text-white' : 'bg-rose-50 text-rose-800 border border-rose-200 hover:bg-rose-100'
                      }`}
                    >
                      Excluded ({excel2ExcludedCount})
                    </button>
                  )}
                </div>

                <div className="text-xs text-slate-500 font-mono">
                  Eligible: <span className="font-bold text-slate-800">{excel2EligibleCount}</span> | Excluded: <span className="font-bold text-slate-800">{excel2ExcludedCount}</span>
                </div>
              </div>

              {/* Table of Calls */}
              {filteredCalls.length === 0 ? (
                <div className="p-8 text-center bg-white rounded-xl border border-slate-200 text-slate-500 text-xs">
                  No inspection call records match the selected filter.
                </div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100/90 text-slate-700 font-semibold border-b border-slate-200">
                          <th className="py-2.5 px-3">Call #</th>
                          <th className="py-2.5 px-3">FOI Number</th>
                          <th className="py-2.5 px-3">Contractor / Main Supplier</th>
                          <th className="py-2.5 px-3">Supplier / Vendor</th>
                          <th className="py-2.5 px-3">Item Description</th>
                          <th className="py-2.5 px-3">Inspection Dates</th>
                          <th className="py-2.5 px-3">Review Dates</th>
                          <th className="py-2.5 px-3">QAP Appl.</th>
                          <th className="py-2.5 px-3 text-right">No. of Days</th>
                          <th className="py-2.5 px-3 text-center">Multiplier</th>
                          <th className="py-2.5 px-3 text-right">Computed MH</th>
                          <th className="py-2.5 px-3">Status / Reason</th>
                          <th className="py-2.5 px-3">Source Cell</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredCalls.map((call, idx) => (
                          <tr key={`call-${call.rowIndex || idx}`} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-2.5 px-3 font-mono font-medium text-slate-900">{call.callNo || `Row ${call.rowIndex}`}</td>
                            <td className="py-2.5 px-3 font-mono text-xs">
                              {call.foiNumber ? (
                                <span className="font-semibold text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-200">
                                  {call.foiNumber}
                                </span>
                              ) : (
                                <span className="text-slate-400 italic text-[11px]">Unavailable</span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-slate-800 font-medium max-w-[140px] truncate" title={call.contractor || 'Not Available'}>
                              {call.contractor || <span className="text-slate-400 italic">Not Available</span>}
                            </td>
                            <td className="py-2.5 px-3 text-slate-800 font-medium max-w-[140px] truncate" title={call.vendor || 'Not Available'}>
                              {call.vendor || <span className="text-slate-400 italic">Not Available</span>}
                            </td>
                            <td className="py-2.5 px-3 text-slate-600 max-w-[140px] truncate" title={call.itemDescription || 'Not Available'}>
                              {call.itemDescription || <span className="text-slate-400 italic">Not Available</span>}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px] text-slate-700">
                              {call.inspectionFrom || call.inspectionUpto ? (
                                <span>{call.inspectionFrom || '—'} → {call.inspectionUpto || '—'}</span>
                              ) : (
                                '—'
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px] text-indigo-700 max-w-[110px] truncate" title={call.reviewDates || '—'}>
                              {call.reviewDates || '—'}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px] text-slate-700">{call.qapApproval || '—'}</td>
                            <td className="py-2.5 px-3 text-right font-mono font-semibold text-slate-900">
                              {call.numberOfDays}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono font-semibold text-slate-600">
                              {call.multiplier}
                            </td>
                            <td className="py-2.5 px-3 text-right">
                              <span className={getValueBadgeClass(call.computedMh)}>
                                {formatDisplayValue(call.computedMh)}
                              </span>
                            </td>
                            <td className="py-2.5 px-3 text-[11px] max-w-[190px]">
                              {call.status === 'INCLUDED_INSPECTION' ? (
                                <span className="text-emerald-800 font-medium">{call.reason}</span>
                              ) : call.status === 'INCLUDED_INSPECTION_WITH_CONFLICT' ? (
                                <span className="text-amber-800 font-medium" title={call.reason}>⚠️ {call.reason}</span>
                              ) : call.status === 'INCLUDED_QAP' ? (
                                <span className="text-purple-800 font-medium">{call.reason}</span>
                              ) : call.status === 'INCLUDED_QAP_WITH_CONFLICT' ? (
                                <span className="text-amber-800 font-medium" title={call.reason}>⚠️ {call.reason}</span>
                              ) : call.status === 'INCLUDED_REVIEW' ? (
                                <span className="text-indigo-800 font-medium">{call.reason}</span>
                              ) : (
                                <span className="text-slate-500">{call.reason || 'Excluded'}</span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px] text-emerald-700 bg-emerald-50/40">
                              {call.sourceCell || `Row ${call.rowIndex}`}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB 4: SOURCE-BY-SOURCE AUDIT TABLE / CELL TRACE */}
          {activeTab === 'lineage' && (
            <div className="space-y-4">
              <div className="p-3 bg-blue-50/60 border border-blue-200 rounded-xl text-xs text-blue-900">
                <strong>Complete Authoritative Cell Trace:</strong> Every rule calculation is traced back to its specific Excel workbook, sheet, row, and cell reference.
              </div>

              {lineage.length === 0 ? (
                <div className="p-8 text-center bg-white rounded-xl border border-slate-200 text-slate-500 text-xs">
                  No cell lineage traces recorded for this job.
                </div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100/90 text-slate-700 font-semibold border-b border-slate-200">
                          <th className="py-2.5 px-3">Field</th>
                          <th className="py-2.5 px-3">Workbook</th>
                          <th className="py-2.5 px-3">Sheet</th>
                          <th className="py-2.5 px-3">Row / Cell</th>
                          <th className="py-2.5 px-3">Column</th>
                          <th className="py-2.5 px-3">Rule Applied</th>
                          <th className="py-2.5 px-3">Calculation Step</th>
                          <th className="py-2.5 px-3 text-right">Result</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {lineage.map((item, idx) => (
                          <tr key={`lin-${idx}`} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-2.5 px-3 font-semibold text-slate-900">{item.field}</td>
                            <td className="py-2.5 px-3 text-slate-700 font-medium">{item.sourceWorkbook}</td>
                            <td className="py-2.5 px-3 text-slate-600 font-mono text-[11px]">{item.sourceSheet}</td>
                            <td className="py-2.5 px-3 font-mono font-semibold text-blue-700 bg-blue-50/40">
                              {item.sourceCell || (item.sourceRow ? `Row ${item.sourceRow}` : 'Header/Formula')}
                            </td>
                            <td className="py-2.5 px-3 text-slate-600 font-mono text-[11px]">{item.column || '—'}</td>
                            <td className="py-2.5 px-3 text-slate-800 text-[11px] max-w-[200px]">{item.ruleApplied}</td>
                            <td className="py-2.5 px-3 font-mono text-[11px] text-slate-700 max-w-[220px]">
                              {item.intermediateCalculation || '—'}
                            </td>
                            <td className="py-2.5 px-3 text-right">
                              <span className={getValueBadgeClass(item.calculatedValue)}>
                                {formatDisplayValue(item.calculatedValue)}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div
          id="evidence-modal-footer"
          className="flex items-center justify-between px-6 py-3.5 border-t border-slate-200 bg-slate-50 shrink-0"
        >
          <div className="text-xs text-slate-500">
            Total Manhours: <strong className="text-slate-900">{formatDisplayValue(total)} MH</strong>
          </div>
          <button
            id="evidence-modal-bottom-close-btn"
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 transition-colors shadow-2xs"
          >
            Close Audit Evidence
          </button>
        </div>
      </div>
    </div>
  );
};
