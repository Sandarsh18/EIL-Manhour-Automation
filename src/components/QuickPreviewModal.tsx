import React, { useState } from 'react';
import { X, FileSpreadsheet, Layers, Columns, Hash } from 'lucide-react';
import { WorkbookMetadata } from '../types/eil.ts';

interface QuickPreviewModalProps {
  metadata: WorkbookMetadata | null;
  onClose: () => void;
}

export const QuickPreviewModal: React.FC<QuickPreviewModalProps> = ({ metadata, onClose }) => {
  if (!metadata) return null;

  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const currentSheet = metadata.sheets[selectedSheetIndex] || metadata.sheets[0];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[85vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50/80">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold text-slate-900">{metadata.fileName}</h3>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-slate-200 text-slate-700">
                  {metadata.type.toUpperCase()}
                </span>
              </div>
              <p className="text-xs text-slate-500">
                Uploaded {new Date(metadata.uploadedAt).toLocaleString()} •{' '}
                {metadata.sheets.length} worksheet(s) detected
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 rounded-md transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Sheet Tabs */}
        {metadata.sheets.length > 1 && (
          <div className="px-6 pt-3 border-b border-slate-200 bg-white flex items-center gap-2 overflow-x-auto">
            <span className="text-xs font-semibold text-slate-400 flex items-center gap-1 shrink-0 mr-1">
              <Layers className="w-3.5 h-3.5" /> Sheets:
            </span>
            {metadata.sheets.map((sheet, idx) => (
              <button
                key={sheet.name}
                onClick={() => setSelectedSheetIndex(idx)}
                className={`px-3 py-1.5 text-xs font-medium rounded-t-md border-b-2 transition whitespace-nowrap ${
                  idx === selectedSheetIndex
                    ? 'border-blue-600 text-blue-600 bg-blue-50/50'
                    : 'border-transparent text-slate-600 hover:text-slate-900'
                }`}
              >
                {sheet.name} ({sheet.rowCount} rows)
              </button>
            ))}
          </div>
        )}

        {/* Sheet Stats Bar */}
        {currentSheet && (
          <div className="px-6 py-2.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4 text-xs">
            <div className="flex items-center gap-4 text-slate-600">
              <span className="flex items-center gap-1.5">
                <Hash className="w-3.5 h-3.5 text-slate-400" />
                <strong>{currentSheet.rowCount}</strong> data rows
              </span>
              <span className="flex items-center gap-1.5">
                <Columns className="w-3.5 h-3.5 text-slate-400" />
                <strong>{currentSheet.columnCount}</strong> detected columns
              </span>
            </div>
            <div className="text-[11px] text-slate-500">
              Showing first {Math.min(10, currentSheet.rowCount)} sample rows
            </div>
          </div>
        )}

        {/* Detected Headers */}
        {currentSheet && currentSheet.headers.length > 0 && (
          <div className="px-6 py-2.5 border-b border-slate-200 bg-white">
            <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1.5">
              Detected Column Headers ({currentSheet.headers.length}):
            </div>
            <div className="flex flex-wrap gap-1.5 max-h-20 overflow-y-auto">
              {currentSheet.headers.map((h, i) => (
                <span
                  key={i}
                  className="px-2 py-0.5 text-xs bg-slate-100 text-slate-700 rounded border border-slate-200/80 font-mono"
                >
                  {h}
                </span>
              ))}
            </div>
          </div>
        )}

        {/* Sample Table */}
        <div className="flex-1 overflow-auto p-6">
          {currentSheet && currentSheet.previewRows.length > 0 ? (
            <div className="border border-slate-200 rounded-lg overflow-hidden shadow-2xs">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                    <tr>
                      <th className="px-3 py-2 border-r border-slate-200 text-slate-400 w-12 text-center">
                        #
                      </th>
                      {currentSheet.headers.map((h, idx) => (
                        <th key={idx} className="px-3 py-2 border-r border-slate-200 whitespace-nowrap">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                    {currentSheet.previewRows.map((row, rIdx) => (
                      <tr key={rIdx} className="hover:bg-blue-50/40 transition">
                        <td className="px-3 py-1.5 border-r border-slate-200 text-slate-400 text-center font-sans">
                          {rIdx + 1}
                        </td>
                        {currentSheet.headers.map((h, cIdx) => (
                          <td
                            key={cIdx}
                            className="px-3 py-1.5 border-r border-slate-200 whitespace-nowrap text-slate-800"
                          >
                            {row[h] !== undefined && row[h] !== null && String(row[h]).trim() !== '' ? (
                              String(row[h])
                            ) : (
                              <span className="text-slate-300 italic">—</span>
                            )}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="text-center py-12 text-slate-400 text-sm">
              No data rows found in this sheet.
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md transition"
          >
            Close Preview
          </button>
        </div>
      </div>
    </div>
  );
};
