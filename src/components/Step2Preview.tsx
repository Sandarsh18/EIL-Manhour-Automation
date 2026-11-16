import React, { useState } from 'react';
import { ArrowLeft, CheckCircle2, AlertTriangle, Table, RefreshCw, Sparkles } from 'lucide-react';
import { WorkbooksState, WorkbookType } from '../types/eil.ts';

interface Step2PreviewProps {
  workbooks: WorkbooksState;
  onBackToStep1: () => void;
}

export const Step2Preview: React.FC<Step2PreviewProps> = ({ workbooks, onBackToStep1 }) => {
  const [activeTab, setActiveTab] = useState<WorkbookType>('excel1');

  const currentWb = workbooks[activeTab];
  const activeSheet = currentWb?.sheets.find((s) => s.name === currentWb.activeSheet) || currentWb?.sheets[0];

  return (
    <div className="space-y-6">
      {/* Header Bar */}
      <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep1}
            className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition"
            title="Return to Step 1 Uploads"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                Step 2 of 6
              </span>
              <h2 className="text-base font-bold text-slate-900">
                Inspect & Map Dynamic Columns
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Review detected worksheet headers and map canonical EIL fields across arbitrary column positions.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onBackToStep1}
            className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md transition"
          >
            ← Back to Step 1 (Uploads)
          </button>
        </div>
      </div>

      {/* Tab Switcher for the 3 Workbooks */}
      <div className="flex items-center gap-2 border-b border-slate-200 bg-white px-4 pt-2 rounded-t-xl">
        {[
          { key: 'excel1' as WorkbookType, label: 'Excel 1 — CONSOLIDATED REPORT' },
          { key: 'excel2' as WorkbookType, label: 'Excel 2 — INSPECTION CALL LOG' },
          { key: 'excel3' as WorkbookType, label: 'Excel 3 — MASTER TEMPLATE' },
        ].map((tab) => (
          <button
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            className={`px-4 py-2 text-xs font-semibold rounded-t-lg border-b-2 transition flex items-center gap-2 ${
              activeTab === tab.key
                ? 'border-blue-600 text-blue-600 bg-blue-50/50'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            <Table className="w-3.5 h-3.5" />
            {tab.label}
          </button>
        ))}
      </div>

      {/* Workbook Inspection & Mapping Container */}
      <div className="bg-white rounded-b-xl border border-slate-200 border-t-0 p-6 space-y-6 shadow-xs">
        {currentWb ? (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-4 p-4 bg-slate-50 rounded-lg border border-slate-200">
              <div>
                <span className="text-xs font-bold text-slate-900 block">{currentWb.fileName}</span>
                <span className="text-[11px] text-slate-500">
                  Active sheet: <strong className="text-slate-700">{activeSheet?.name}</strong> •{' '}
                  {activeSheet?.rowCount} rows • {activeSheet?.columnCount} columns
                </span>
              </div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-1 text-xs font-semibold bg-emerald-50 text-emerald-700 rounded border border-emerald-200 flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  Headers Recognized
                </span>
              </div>
            </div>

            {/* Detected Headers */}
            <div>
              <h4 className="text-xs font-bold text-slate-800 mb-2">Detected Sheet Headers:</h4>
              <div className="flex flex-wrap gap-1.5">
                {activeSheet?.headers.map((h, i) => (
                  <span
                    key={i}
                    className="px-2.5 py-1 text-xs bg-slate-100 text-slate-700 rounded border border-slate-200 font-mono"
                  >
                    {h}
                  </span>
                ))}
              </div>
            </div>

            {/* Sample Rows Preview */}
            <div>
              <h4 className="text-xs font-bold text-slate-800 mb-2">Sample Source Rows (First 5):</h4>
              <div className="border border-slate-200 rounded-lg overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 border-b border-slate-200">
                    <tr>
                      <th className="px-3 py-2 text-slate-400 w-10">#</th>
                      {activeSheet?.headers.map((h, i) => (
                        <th key={i} className="px-3 py-2 whitespace-nowrap">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                    {activeSheet?.previewRows.slice(0, 5).map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50">
                        <td className="px-3 py-1.5 text-slate-400">{idx + 1}</td>
                        {activeSheet?.headers.map((h, cIdx) => (
                          <td key={cIdx} className="px-3 py-1.5 whitespace-nowrap text-slate-800">
                            {row[h] !== undefined && row[h] !== null ? String(row[h]) : '—'}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="text-center py-12 text-slate-400 text-xs">
            Workbook data not found. Please upload in Step 1.
          </div>
        )}
      </div>
    </div>
  );
};
