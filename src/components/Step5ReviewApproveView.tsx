import React, { useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  AlertTriangle,
  AlertCircle,
  Edit2,
  CheckCheck,
  Filter,
  Save,
  Clock,
  Sparkles,
  FileText,
  Trash2,
  RotateCcw,
  ExternalLink,
  ShieldCheck,
  Database,
  Calendar,
  Layers,
  ChevronDown,
  ChevronRight,
  Info,
  X,
  RefreshCw,
} from 'lucide-react';
import { ReviewItem, CalculationLineageItem, Excel1OrderRecord, Excel2InspectionRecord, WarningAuditDetail } from '../types/eil.ts';
import { isPresentValue, formatDisplayValue, getValueBadgeClass } from '../utils/eilHighlight';
import { CalculationEvidenceModal } from './CalculationEvidenceModal';

interface Step5ReviewApproveViewProps {
  reviewItems: Record<string, ReviewItem>;
  onUpdateReviewItem: (jobNumber: string, updates: Partial<ReviewItem>) => Promise<void>;
  onApproveAll: () => Promise<void>;
  onUnapproveAll: () => Promise<void>;
  onDeleteAll: () => Promise<void>;
  onDeleteReviewItem: (jobNumber: string) => Promise<void>;
  onBackToStep4: () => void;
  onProceedToStep6: () => void;
  isGenerating?: boolean;
}

export const Step5ReviewApproveView: React.FC<Step5ReviewApproveViewProps> = ({
  reviewItems,
  onUpdateReviewItem,
  onApproveAll,
  onUnapproveAll,
  onDeleteAll,
  onDeleteReviewItem,
  onBackToStep4,
  onProceedToStep6,
  isGenerating = false,
}) => {
  const [editingJob, setEditingJob] = useState<ReviewItem | null>(null);
  const [auditDetailJob, setAuditDetailJob] = useState<ReviewItem | null>(null);
  const [activeAuditTab, setActiveAuditTab] = useState<'summary' | 'excel1' | 'excel2' | 'lineage'>('summary');
  const [filterStatus, setFilterStatus] = useState<'ALL' | 'APPROVED' | 'OVERRIDDEN'>('ALL');
  const [isSaving, setIsSaving] = useState(false);
  const [confirmingDeleteJob, setConfirmingDeleteJob] = useState<string | null>(null);
  const [confirmingDeleteAll, setConfirmingDeleteAll] = useState(false);

  const items = ((Object.values(reviewItems) || []) as ReviewItem[]).filter(
    (i) => i && i.jobNumber && i.jobNumber.trim() !== ''
  );
  const totalCount = items.length;
  const approvedCount = items.filter((i) => i.approvalStatus === 'APPROVED').length;
  const pendingCount = items.filter((i) => i.approvalStatus !== 'APPROVED').length;

  const filteredItems = items.filter((item) => {
    if (filterStatus === 'APPROVED') return item.approvalStatus === 'APPROVED';
    if (filterStatus === 'OVERRIDDEN') return item.isOverridden;
    return true;
  });

  const handleSaveEdit = async () => {
    if (!editingJob) return;
    try {
      setIsSaving(true);
      await onUpdateReviewItem(editingJob.jobNumber, {
        fd: editingJob.fd,
        runningOrders: editingJob.runningOrders,
        ocsDone: editingJob.ocsDone,
        expediting: editingJob.expediting,
        inspection: editingJob.inspection,
        others: editingJob.others,
        meeting: editingJob.meeting,
        auditNotes: editingJob.auditNotes,
        approvalStatus: editingJob.approvalStatus,
        isOverridden: true,
      });
      setEditingJob(null);
    } catch (err: any) {
      console.error('Failed to update review item:', err);
    } finally {
      setIsSaving(false);
    }
  };

  const handleQuickApprove = async (jobNumber: string) => {
    await onUpdateReviewItem(jobNumber, { approvalStatus: 'APPROVED' });
  };

  const handleQuickUnapprove = async (jobNumber: string) => {
    await onUpdateReviewItem(jobNumber, { approvalStatus: 'PENDING' });
  };

  const openAuditInspector = (item: ReviewItem, tab: 'summary' | 'excel1' | 'excel2' | 'lineage' = 'summary') => {
    setAuditDetailJob(item);
    setActiveAuditTab(tab);
  };

  return (
    <div className="space-y-6 pb-24">
      {/* Header Bar */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep4}
            className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition"
            title="Return to Step 4 Rule Analysis"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                Step 5 of 6
              </span>
              <h2 className="text-base font-bold text-slate-900">
                Audit Review & Verification
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Inspect calculated hours, perform authorized overrides with audit notes, and approve records for final reporting.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            id="btn-approve-all"
            onClick={onApproveAll}
            className="px-3.5 py-2 text-xs font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-lg transition flex items-center gap-1.5 shadow-2xs"
            title="Mark all non-blocked jobs as Approved"
          >
            <CheckCheck className="w-4 h-4 text-emerald-600" />
            <span>Approve All</span>
          </button>

          <button
            id="btn-unapprove-all"
            onClick={onUnapproveAll}
            className="px-3.5 py-2 text-xs font-bold text-amber-800 bg-amber-50 hover:bg-amber-100 border border-amber-300 rounded-lg transition flex items-center gap-1.5 shadow-2xs"
            title="Reset all jobs to Pending"
          >
            <RotateCcw className="w-3.5 h-3.5 text-amber-600" />
            <span>Unapprove All</span>
          </button>

          {confirmingDeleteAll ? (
            <div className="flex items-center gap-1.5 bg-rose-50 border border-rose-200 px-3 py-1.5 rounded-lg text-xs">
              <span className="text-[11px] text-rose-800 font-semibold">Delete all {totalCount} jobs?</span>
              <button
                onClick={async () => {
                  setConfirmingDeleteAll(false);
                  await onDeleteAll();
                }}
                className="px-2 py-0.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-[10px] font-bold"
              >
                Confirm
              </button>
              <button
                onClick={() => setConfirmingDeleteAll(false)}
                className="px-1.5 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-[10px]"
              >
                Cancel
              </button>
            </div>
          ) : (
            <button
              id="btn-delete-all"
              onClick={() => setConfirmingDeleteAll(true)}
              className="px-3.5 py-2 text-xs font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-300 rounded-lg transition flex items-center gap-1.5 shadow-2xs"
              title="Remove all review jobs from batch"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-600" />
              <span>Delete All</span>
            </button>
          )}
        </div>
      </div>

      {/* Summary KPI Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="bg-white rounded-xl p-3.5 border border-slate-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-slate-500 uppercase">Total in Scope</span>
          <div className="text-xl font-bold text-slate-900 mt-1">{totalCount}</div>
          <span className="text-[10px] text-slate-400">Total job records</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-emerald-200 bg-emerald-50/20 shadow-2xs">
          <span className="text-[11px] font-semibold text-emerald-800 uppercase flex items-center gap-1">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
            Approved
          </span>
          <div className="text-xl font-bold text-emerald-800 mt-1">{approvedCount}</div>
          <span className="text-[10px] text-emerald-600">Ready for Excel 3 output</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-slate-200 bg-slate-50/50 shadow-2xs">
          <span className="text-[11px] font-semibold text-slate-700 uppercase flex items-center gap-1">
            <Clock className="w-3.5 h-3.5 text-slate-500" />
            Not Approved
          </span>
          <div className="text-xl font-bold text-slate-800 mt-1">{pendingCount}</div>
          <span className="text-[10px] text-slate-500">Awaiting auditor review</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-indigo-200 bg-indigo-50/20 shadow-2xs">
          <span className="text-[11px] font-semibold text-indigo-800 uppercase flex items-center gap-1">
            <Sparkles className="w-3.5 h-3.5 text-indigo-600" />
            Auditor Overrides
          </span>
          <div className="text-xl font-bold text-indigo-900 mt-1">
            {items.filter((i) => i.isOverridden).length}
          </div>
          <span className="text-[10px] text-indigo-600">Manual adjustments with notes</span>
        </div>
      </div>

      {/* Review Table Card */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        {/* Filters */}
        <div className="p-3.5 border-b border-slate-200 bg-slate-50 flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-slate-500" />
            <span className="text-xs font-semibold text-slate-700">Filter Review List:</span>
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

            <div className="flex items-center gap-1 text-xs">
              {(['ALL', 'APPROVED', 'OVERRIDDEN'] as const).map((tab) => (
                <button
                  key={tab}
                  onClick={() => setFilterStatus(tab)}
                  className={`px-3 py-1 rounded-md transition font-medium ${
                    filterStatus === tab
                      ? 'bg-blue-600 text-white shadow-2xs'
                      : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
                  }`}
                >
                  {tab === 'ALL' ? 'All' : tab === 'APPROVED' ? 'Approved' : 'Overridden'}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* Table */}
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
                <th className="px-3 py-2.5 text-right bg-blue-50/50 font-bold text-blue-900">Total (h)</th>
                <th className="px-3 py-2.5 text-center">Status</th>
                <th className="px-3 py-2.5 text-center">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
              {filteredItems.length === 0 ? (
                <tr>
                  <td colSpan={13} className="px-4 py-12 text-center text-slate-400 font-sans">
                    No review items match the selected filter.
                  </td>
                </tr>
              ) : (
                filteredItems.map((item, idx) => (
                  <tr
                    key={`review-${item.jobNumber || 'item'}-${idx}`}
                    id={`review-row-${item.jobNumber || idx}`}
                    className="hover:bg-slate-50/80 transition"
                  >
                    <td className="px-3 py-2.5 font-bold text-slate-900">
                      <div className="flex items-center gap-1.5">
                        <span>{item.jobNumber}</span>
                        {item.isOverridden && (
                          <span className="text-[9px] bg-indigo-100 text-indigo-700 px-1 py-0.2 rounded font-sans font-bold">
                            OVERRIDDEN
                          </span>
                        )}
                      </div>
                    </td>
                    <td className="px-3 py-2.5 font-mono text-slate-600 text-xs">
                      {item.foiNumber ? (
                        <span className="font-semibold text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-200">
                          {item.foiNumber}
                        </span>
                      ) : (
                        <span className="text-slate-400 italic">Unavailable</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-center font-sans">
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
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.fd)}>
                        {formatDisplayValue(item.fd)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.runningOrders)}>
                        {formatDisplayValue(item.runningOrders)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.ocsDone)}>
                        {formatDisplayValue(item.ocsDone)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.expediting)}>
                        {formatDisplayValue(item.expediting)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.inspection)}>
                        {formatDisplayValue(item.inspection)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.others)}>
                        {formatDisplayValue(item.others)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right">
                      <span className={getValueBadgeClass(item.meeting)}>
                        {formatDisplayValue(item.meeting)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right bg-blue-50/30">
                      <span className={getValueBadgeClass(item.total, true)}>
                        {formatDisplayValue(item.total)}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-center font-sans">
                      {item.approvalStatus === 'APPROVED' ? (
                        <button
                          onClick={() => openAuditInspector(item, 'summary')}
                          className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-700 bg-emerald-100 hover:bg-emerald-200 px-2 py-0.5 rounded-full transition"
                          title="Click to view calculation evidence"
                        >
                          <CheckCircle2 className="w-3 h-3" />
                          Approved
                        </button>
                      ) : (
                        <button
                          onClick={() => openAuditInspector(item, 'summary')}
                          className="inline-flex items-center gap-1 text-[10px] font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 px-2 py-0.5 rounded-full transition"
                          title="Click to view calculation evidence"
                        >
                          Not Approved
                        </button>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-center font-sans">
                      <div className="flex items-center justify-center gap-1.5">
                        {item.approvalStatus === 'APPROVED' ? (
                          <button
                            onClick={() => handleQuickUnapprove(item.jobNumber)}
                            className="px-2 py-1 text-[10px] font-semibold text-amber-700 bg-amber-50 hover:bg-amber-100 border border-amber-200 rounded transition flex items-center gap-1"
                            title="Reset to Pending"
                          >
                            <Clock className="w-3 h-3" />
                            <span>Unapprove</span>
                          </button>
                        ) : (
                          <button
                            onClick={() => handleQuickApprove(item.jobNumber)}
                            className="px-2 py-1 text-[10px] font-semibold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded transition flex items-center gap-1"
                            title="Approve Job"
                          >
                            <CheckCircle2 className="w-3 h-3" />
                            <span>Approve</span>
                          </button>
                        )}

                        <button
                          onClick={() => openAuditInspector(item, 'summary')}
                          className="p-1 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded transition flex items-center gap-0.5"
                          title="View Full Audit Trail, Source Orders & Inspection Logs"
                        >
                          <FileText className="w-3.5 h-3.5" />
                        </button>

                        <button
                          onClick={() => setEditingJob({ ...item })}
                          className="p-1 text-slate-500 hover:text-blue-600 hover:bg-blue-50 rounded transition"
                          title="Edit Hours & Add Audit Notes"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>

                        {confirmingDeleteJob === item.jobNumber ? (
                          <div className="inline-flex items-center gap-1.5 bg-rose-50 border border-rose-300 px-2 py-0.5 rounded text-[11px] shadow-xs">
                            <span className="text-rose-800 font-medium">Delete job {item.jobNumber}?</span>
                            <button
                              onClick={() => setConfirmingDeleteJob(null)}
                              className="px-1.5 py-0.5 bg-slate-200 text-slate-700 rounded hover:bg-slate-300 font-medium text-[10px]"
                            >
                              Cancel
                            </button>
                            <button
                              onClick={async () => {
                                const jn = item.jobNumber;
                                setConfirmingDeleteJob(null);
                                if (jn) await onDeleteReviewItem(jn);
                              }}
                              className="px-1.5 py-0.5 bg-rose-600 text-white rounded font-bold hover:bg-rose-700 text-[10px]"
                            >
                              Delete
                            </button>
                          </div>
                        ) : (
                          <button
                            onClick={() => setConfirmingDeleteJob(item.jobNumber)}
                            className="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition"
                            title={`Delete job ${item.jobNumber}`}
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
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
            Approval Progress: {approvedCount} of {totalCount} jobs approved
          </span>
          {pendingCount > 0 && (
            <span className="text-xs text-amber-600">
              ({pendingCount} pending, you can still generate output or click "Approve All")
            </span>
          )}
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep4}
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition"
          >
            ← Back to Step 4
          </button>

          <button
            id="btn-proceed-step6"
            onClick={onProceedToStep6}
            disabled={isGenerating || (approvedCount === 0 && totalCount > 0)}
            className={`px-5 py-2.5 text-xs font-bold rounded-xl transition flex items-center gap-2 shadow-xs ${
              isGenerating
                ? 'bg-blue-400 text-white cursor-wait'
                : approvedCount > 0 || totalCount === 0
                ? 'bg-blue-600 hover:bg-blue-700 text-white cursor-pointer shadow-blue-500/20'
                : 'bg-slate-200 text-slate-400 border border-slate-300 cursor-not-allowed'
            }`}
          >
            {isGenerating ? (
              <>
                <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                <span>Generating Output...</span>
              </>
            ) : (
              <>
                <span>Generate Output →</span>
                <ArrowRight className="w-4 h-4" />
              </>
            )}
          </button>
        </div>
      </div>

      {/* Manual Override & Audit Notes Edit Modal */}
      {editingJob && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-xl w-full p-6 space-y-5">
            <div className="flex items-start justify-between border-b border-slate-100 pb-3">
              <div>
                <span className="text-[10px] font-bold uppercase text-blue-600">
                  Manual Override & Audit Notes
                </span>
                <h3 className="text-base font-bold text-slate-900 font-mono mt-0.5">
                  Job Number: {editingJob.jobNumber}
                </h3>
              </div>
              <button
                onClick={() => setEditingJob(null)}
                className="text-slate-400 hover:text-slate-600 font-bold text-lg"
              >
                ✕
              </button>
            </div>

            <div className="space-y-4 text-xs font-sans">
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 font-mono">
                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    FD:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.fd}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, fd: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Running Orders:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.runningOrders}
                    onChange={(e) =>
                      setEditingJob({
                        ...editingJob,
                        runningOrders: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    OCS Done:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.ocsDone}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, ocsDone: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Expediting:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.expediting}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, expediting: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Inspection:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.inspection}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, inspection: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Others:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.others}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, others: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Meeting:
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={editingJob.meeting}
                    onChange={(e) =>
                      setEditingJob({ ...editingJob, meeting: parseFloat(e.target.value) || 0 })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-sans font-semibold text-slate-600 mb-1">
                    Approval Status:
                  </label>
                  <select
                    value={editingJob.approvalStatus}
                    onChange={(e) =>
                      setEditingJob({
                        ...editingJob,
                        approvalStatus: e.target.value as any,
                      })
                    }
                    className="w-full text-xs border border-slate-300 rounded px-2 py-1.5"
                  >
                    <option value="PENDING">PENDING</option>
                    <option value="APPROVED">APPROVED</option>
                    <option value="FLAGGED">FLAGGED</option>
                    <option value="BLOCKED">BLOCKED</option>
                  </select>
                </div>
              </div>

              {/* Live Recalculation Preview Banner */}
              <div className="p-3 bg-blue-50/80 border border-blue-200 rounded-lg flex items-center justify-between">
                <div>
                  <span className="text-[10px] font-bold uppercase text-blue-700 tracking-wider block">Live Formula Recalculation</span>
                  <span className="text-[11px] text-blue-900 font-mono">
                    Total = Exp ({editingJob.expediting || 0}) + Insp ({editingJob.inspection || 0}) + Others ({editingJob.others || 0}) + Meeting ({editingJob.meeting || 0})
                  </span>
                </div>
                <div className="font-mono text-base font-bold text-blue-900 bg-white px-2.5 py-1 rounded border border-blue-200">
                  {((Number(editingJob.expediting) || 0) + (Number(editingJob.inspection) || 0) + (Number(editingJob.others) || 0) + (Number(editingJob.meeting) || 0)).toFixed(1)} h
                </div>
              </div>

              {/* Audit Notes field */}
              <div>
                <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                  Auditor Justification / Notes (Required for Overrides):
                </label>
                <textarea
                  rows={3}
                  value={editingJob.auditNotes}
                  onChange={(e) => setEditingJob({ ...editingJob, auditNotes: e.target.value })}
                  placeholder="e.g., Authorized adjustment per Regional Manager instruction due to offshore call delay..."
                  className="w-full text-xs border border-slate-300 rounded-lg p-2.5 text-slate-800"
                />
              </div>
            </div>

            <div className="pt-3 border-t border-slate-100 flex items-center justify-end gap-2">
              <button
                onClick={() => setEditingJob(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg transition"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveEdit}
                disabled={isSaving}
                className="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition flex items-center gap-1.5 shadow-2xs"
              >
                <Save className="w-3.5 h-3.5" />
                <span>Save Override</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Authoritative Calculation Evidence & Lineage Modal */}
      {auditDetailJob && (
        <CalculationEvidenceModal
          job={auditDetailJob}
          initialTab={activeAuditTab}
          onEditOverride={(target) => {
            setAuditDetailJob(null);
            setEditingJob({ ...(target as ReviewItem) });
          }}
          onClose={() => setAuditDetailJob(null)}
        />
      )}
    </div>
  );
};
