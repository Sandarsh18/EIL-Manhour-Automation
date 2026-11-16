import React, { useState, useMemo } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  Search,
  CheckCircle2,
  AlertTriangle,
  Layers,
  FileSpreadsheet,
  PhoneCall,
  Eye,
} from 'lucide-react';
import { JobReconciliationSummary, ReconciledJob } from '../types/eil.ts';

interface Step3JobSelectViewProps {
  reconciliation: JobReconciliationSummary | null;
  selectedJobs: string[];
  onSelectedJobsChange: (jobs: string[]) => void;
  onBackToStep2: () => void;
  onProceedToStep4: () => void;
}

export const Step3JobSelectView: React.FC<Step3JobSelectViewProps> = ({
  reconciliation,
  selectedJobs,
  onSelectedJobsChange,
  onBackToStep2,
  onProceedToStep4,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [filterTab, setFilterTab] = useState<'ALL' | 'EXCEL1' | 'EXCEL2_ONLY' | 'BOTH'>('ALL');
  const [inspectingJob, setInspectingJob] = useState<ReconciledJob | null>(null);

  const jobs = (reconciliation?.jobs || []).filter(
    (j) => j && j.jobNumber && j.jobNumber.trim() !== ''
  );

  const filteredJobs = useMemo(() => {
    return jobs.filter((job) => {
      // Search filter (searches job number and FOI number)
      const q = searchQuery.trim().toUpperCase();
      const matchesSearch =
        !q ||
        job.jobNumber.toUpperCase().includes(q) ||
        (job.foiNumber && job.foiNumber.toUpperCase().includes(q));

      if (!matchesSearch) return false;

      // Tab filter
      if (filterTab === 'EXCEL1') return job.inExcel1;
      if (filterTab === 'EXCEL2_ONLY') return !job.inExcel1 && job.inExcel2;
      if (filterTab === 'BOTH') return job.inExcel1 && job.inExcel2;
      return true;
    });
  }, [jobs, searchQuery, filterTab]);

  const handleToggleJob = (jn: string) => {
    if (selectedJobs.includes(jn)) {
      onSelectedJobsChange(selectedJobs.filter((j) => j !== jn));
    } else {
      onSelectedJobsChange([...selectedJobs, jn]);
    }
  };

  const handleSelectAll = () => {
    const allFiltered = filteredJobs.map((j) => j.jobNumber);
    const merged = Array.from(new Set([...selectedJobs, ...allFiltered]));
    onSelectedJobsChange(merged);
  };

  const handleClearAll = () => {
    const filteredSet = new Set(filteredJobs.map((j) => j.jobNumber));
    onSelectedJobsChange(selectedJobs.filter((j) => !filteredSet.has(j)));
  };

  const handleSelectWithCalls = () => {
    const withCalls = jobs.filter((j) => j.inExcel2).map((j) => j.jobNumber);
    onSelectedJobsChange(Array.from(new Set(withCalls)));
  };

  const excel1Count = reconciliation?.excel1Count ?? jobs.filter((j) => j.inExcel1).length;
  const excel2OnlyCount = reconciliation?.excel2OnlyCount ?? jobs.filter((j) => !j.inExcel1 && j.inExcel2).length;
  const bothCount = jobs.filter((j) => j.inExcel1 && j.inExcel2).length;
  const matchedAllCount = reconciliation?.matchedAll ?? jobs.filter((j) => j.inExcel1 && j.inExcel2 && j.inExcel3).length;

  return (
    <div className="space-y-6 pb-24">
      {/* Header Bar */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep2}
            className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition"
            title="Return to Step 2 Mappings"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                Step 3 of 6
              </span>
              <h2 className="text-base font-bold text-slate-900">
                Job Discovery & Reconciliation (J1 ∪ J2)
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Eligible jobs are discovered from both Excel 1 (Consolidated Report) and Excel 2 (PM Call Logs with current-month activity).
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onBackToStep2}
            className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md transition"
          >
            ← Back to Step 2
          </button>
        </div>
      </div>

      {/* Authoritative Directive Banner */}
      <div className="bg-blue-50/80 border border-blue-200 rounded-xl p-4 flex items-start gap-3">
        <CheckCircle2 className="w-5 h-5 text-blue-600 shrink-0 mt-0.5" />
        <div className="text-xs text-blue-950 leading-relaxed">
          <span className="font-bold text-blue-900">Authoritative Job Universe Rule:</span> The final job list comprises <span className="font-bold">Final Jobs = J1 ∪ J2</span> ({jobs.length} total jobs). Jobs in Excel 1 are evaluated for orders. Jobs found in PM Call Logs (Excel 2) with valid current-month activity are included even if absent from Excel 1 (Case C), while historical-only Excel 2 jobs are excluded (Case D).
        </div>
      </div>

      {/* Reconciliation KPI Metric Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="bg-white rounded-xl p-3.5 border border-slate-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-slate-500 uppercase flex items-center gap-1">
            <Layers className="w-3.5 h-3.5 text-slate-400" />
            Total Eligible Universe
          </span>
          <div className="text-xl font-bold text-slate-900 mt-1">{jobs.length}</div>
          <span className="text-[10px] text-slate-400">J1 ∪ J2 active jobs</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-blue-200 bg-blue-50/30 shadow-2xs">
          <span className="text-[11px] font-semibold text-blue-800 uppercase flex items-center gap-1">
            <FileSpreadsheet className="w-3.5 h-3.5 text-blue-600" />
            Excel 1 Orders (J1)
          </span>
          <div className="text-xl font-bold text-blue-800 mt-1">{excel1Count}</div>
          <span className="text-[10px] text-blue-600">Present in Consolidated Report</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-indigo-200 bg-indigo-50/30 shadow-2xs">
          <span className="text-[11px] font-semibold text-indigo-800 uppercase flex items-center gap-1">
            <PhoneCall className="w-3.5 h-3.5 text-indigo-600" />
            Excel 2 Only (Case C)
          </span>
          <div className="text-xl font-bold text-indigo-800 mt-1">{excel2OnlyCount}</div>
          <span className="text-[10px] text-indigo-600">Discovered from current month calls</span>
        </div>

        <div className="bg-white rounded-xl p-3.5 border border-emerald-200 bg-emerald-50/30 shadow-2xs">
          <span className="text-[11px] font-semibold text-emerald-800 uppercase flex items-center gap-1">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
            In All 3 Workbooks
          </span>
          <div className="text-xl font-bold text-emerald-800 mt-1">{matchedAllCount}</div>
          <span className="text-[10px] text-emerald-600">Full tripartite match</span>
        </div>
      </div>

      {/* Main Table Card */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        {/* Filter and Search Bar */}
        <div className="p-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 bg-slate-50/60">
          <div className="flex items-center gap-2 flex-1 min-w-[240px] max-w-md">
            <div className="relative w-full">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
              <input
                id="search-job-input"
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search Job No or FOI No (e.g. B269, FOI-2691, G660)..."
                className="w-full text-xs bg-white border border-slate-300 rounded-lg pl-9 pr-3 py-2 text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-mono"
              />
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {/* Filter Tabs */}
            <div className="flex items-center bg-slate-200/70 p-0.5 rounded-lg text-xs font-medium">
              <button
                onClick={() => setFilterTab('ALL')}
                className={`px-3 py-1 rounded-md transition ${
                  filterTab === 'ALL' ? 'bg-white text-slate-900 shadow-2xs font-bold' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                All ({jobs.length})
              </button>
              <button
                onClick={() => setFilterTab('EXCEL1')}
                className={`px-3 py-1 rounded-md transition ${
                  filterTab === 'EXCEL1' ? 'bg-white text-slate-900 shadow-2xs font-bold' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Consolidated ({excel1Count})
              </button>
              <button
                onClick={() => setFilterTab('EXCEL2_ONLY')}
                className={`px-3 py-1 rounded-md transition ${
                  filterTab === 'EXCEL2_ONLY' ? 'bg-white text-slate-900 shadow-2xs font-bold' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                PM Calls Only ({excel2OnlyCount})
              </button>
              <button
                onClick={() => setFilterTab('BOTH')}
                className={`px-3 py-1 rounded-md transition ${
                  filterTab === 'BOTH' ? 'bg-white text-slate-900 shadow-2xs font-bold' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                In Both ({bothCount})
              </button>
            </div>

            {/* Select/Clear All */}
            <div className="flex items-center gap-1.5 pl-2 border-l border-slate-300">
              <button
                id="btn-select-with-calls"
                onClick={handleSelectWithCalls}
                className="px-2.5 py-1 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-md transition"
                title="Select jobs with active calls in Excel 2"
              >
                With Calls ({jobs.filter((j) => j.inExcel2).length})
              </button>
              <button
                id="btn-select-all"
                onClick={handleSelectAll}
                className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md transition"
              >
                Select All
              </button>
              <button
                id="btn-clear-all"
                onClick={handleClearAll}
                className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md transition"
              >
                Clear All
              </button>
            </div>
          </div>
        </div>

        {/* Jobs Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 uppercase font-semibold">
              <tr>
                <th className="px-4 py-3 w-12 text-center">Select</th>
                <th className="px-4 py-3">Job Number</th>
                <th className="px-4 py-3">FOI Number</th>
                <th className="px-4 py-3 text-center">Origin</th>
                <th className="px-4 py-3 text-center">Excel 1 (Consol.)</th>
                <th className="px-4 py-3 text-center">Excel 2 (PM Calls)</th>
                <th className="px-4 py-3 text-center">Excel 3 (Template)</th>
                <th className="px-4 py-3 text-right">Details</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-sans">
              {filteredJobs.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-12 text-center text-slate-400">
                    No jobs matching current query or filters.
                  </td>
                </tr>
              ) : (
                filteredJobs.map((job, idx) => {
                  const isSelected = selectedJobs.includes(job.jobNumber);
                  return (
                    <tr
                      key={`select-job-${job.jobNumber || 'job'}-${idx}`}
                      id={`job-row-${job.jobNumber || idx}`}
                      className={`hover:bg-slate-50 transition cursor-pointer ${
                        isSelected ? 'bg-blue-50/30' : ''
                      }`}
                      onClick={() => handleToggleJob(job.jobNumber)}
                    >
                      <td className="px-4 py-3 text-center" onClick={(e) => e.stopPropagation()}>
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => handleToggleJob(job.jobNumber)}
                          className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4 cursor-pointer"
                        />
                      </td>
                      <td className="px-4 py-3 font-mono font-bold text-slate-900 text-xs">
                        {job.jobNumber}
                      </td>
                      <td className="px-4 py-3 font-mono text-slate-700 text-xs">
                        {job.foiNumber ? (
                          <span className="font-semibold text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">
                            {job.foiNumber}
                          </span>
                        ) : (
                          <span className="text-slate-400 italic">Unavailable</span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-center">
                        <span
                          className={`inline-flex items-center text-[10px] font-semibold px-2 py-0.5 rounded-full ${
                            job.origin === 'Both'
                              ? 'bg-emerald-100 text-emerald-800'
                              : job.origin === 'Consolidated Report'
                              ? 'bg-blue-100 text-blue-800'
                              : 'bg-purple-100 text-purple-800'
                          }`}
                        >
                          {job.origin || (job.inExcel1 && job.inExcel2 ? 'Both' : job.inExcel1 ? 'Consolidated Report' : 'PM Call Logs')}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-center">
                        {job.inExcel1 ? (
                          <span className="inline-flex items-center gap-1 text-[11px] font-medium text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                            {job.excel1RowCount} orders
                          </span>
                        ) : (
                          <span className="text-[11px] text-slate-400">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-center">
                        {job.inExcel2 ? (
                          <div className="flex flex-col items-center gap-0.5">
                            <span className="inline-flex items-center gap-1 text-[11px] font-medium text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                              {job.excel2RowCount} calls ({job.excel2EligibleRowCount || 0} active)
                            </span>
                            {job.excel2EligibilityStatus && (
                              <span
                                className={`text-[9px] font-medium px-1.5 py-0.2 rounded ${
                                  job.excel2EligibilityStatus === 'ELIGIBLE_IN_MONTH'
                                    ? 'text-emerald-700 bg-emerald-50'
                                    : 'text-amber-700 bg-amber-50'
                                }`}
                              >
                                {job.excel2EligibilityStatus === 'ELIGIBLE_IN_MONTH'
                                  ? 'Month Active'
                                  : 'Outside Month'}
                              </span>
                            )}
                          </div>
                        ) : (
                          <span className="text-[11px] text-slate-400">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-center">
                        {job.inExcel3 ? (
                          <span className="inline-flex items-center gap-1 text-[11px] font-medium text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                            Present
                          </span>
                        ) : (
                          <span className="text-[11px] text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded border border-amber-200">
                            Missing
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-right" onClick={(e) => e.stopPropagation()}>
                        <button
                          onClick={() => setInspectingJob(job)}
                          className="p-1 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition"
                          title="View cross-workbook presence"
                        >
                          <Eye className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sticky Action Footer */}
      <div className="sticky bottom-0 z-40 bg-white/95 backdrop-blur-md rounded-xl p-4 border border-slate-200 shadow-lg flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <span className="text-xs font-bold text-slate-800">
            {selectedJobs.length} of {jobs.length} jobs selected for rule evaluation
          </span>
          {selectedJobs.length === 0 && (
            <span className="text-xs text-rose-600 font-medium">
              (Please select at least one job number to continue)
            </span>
          )}
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep2}
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition"
          >
            ← Back to Step 2
          </button>

          <button
            id="btn-proceed-step4"
            onClick={onProceedToStep4}
            disabled={selectedJobs.length === 0}
            className={`px-5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-2 shadow-xs ${
              selectedJobs.length > 0
                ? 'bg-blue-600 hover:bg-blue-700 text-white cursor-pointer'
                : 'bg-slate-200 text-slate-400 border border-slate-300 cursor-not-allowed'
            }`}
          >
            <span>Proceed to Step 4: Rule Analysis →</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Inspect Single Job Modal */}
      {inspectingJob && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-lg w-full p-6 space-y-4">
            <div className="flex items-start justify-between border-b border-slate-100 pb-3">
              <div>
                <span className="text-[10px] font-bold uppercase text-slate-400">Reconciliation Detail</span>
                <h3 className="text-base font-bold text-slate-900 font-mono mt-0.5">
                  Job Number: {inspectingJob.jobNumber}
                </h3>
              </div>
              <button
                onClick={() => setInspectingJob(null)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold"
              >
                ✕
              </button>
            </div>

            <div className="space-y-2 text-xs text-slate-700">
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="font-semibold text-slate-500">FOI Number:</span>
                <span className="text-slate-900 font-mono font-medium">
                  {inspectingJob.foiNumber || 'Unavailable in source records'}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="font-semibold text-slate-500">Discovery Origin:</span>
                <span className="text-slate-900 font-medium">
                  {inspectingJob.origin || 'Consolidated Report'}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="font-semibold text-slate-500">Excel 1 Records:</span>
                <span className="text-slate-900 font-mono font-medium">
                  {inspectingJob.inExcel1 ? `${inspectingJob.excel1RowCount} line items` : 'Not Found (Case C Job)'}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="font-semibold text-slate-500">Excel 2 PM Calls:</span>
                <span className="text-slate-900 font-mono font-medium">
                  {inspectingJob.inExcel2
                    ? `${inspectingJob.excel2RowCount} calls logged (${inspectingJob.excel2EligibleRowCount || 0} in active month)`
                    : 'Not Found'}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="font-semibold text-slate-500">Excel 3 Master Template:</span>
                <span className="text-slate-900 font-mono font-medium">
                  {inspectingJob.inExcel3 ? 'Present in template rows' : 'Missing from template (will be inserted)'}
                </span>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                onClick={() => setInspectingJob(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
