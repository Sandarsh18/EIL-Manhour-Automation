import React from 'react';
import { Layers, FolderOpen, Save, RotateCcw, Calendar, CheckCircle2 } from 'lucide-react';

interface HeaderProps {
  evaluationMonth: string;
  onEvaluationMonthChange: (month: string) => void;
  onOpenSaveModal: () => void;
  onOpenLoadModal: () => void;
  onReset: () => void;
  hasWorkbooks: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  evaluationMonth,
  onEvaluationMonthChange,
  onOpenSaveModal,
  onOpenLoadModal,
  onReset,
  hasWorkbooks,
}) => {
  // Compute OCS 6-calendar-month window for display
  const getOcsPeriod = (evalMonth: string) => {
    try {
      const [yearStr, monthStr] = evalMonth.split('-');
      const y = parseInt(yearStr, 10);
      const m = parseInt(monthStr, 10);
      if (isNaN(y) || isNaN(m)) return '';

      const lastDay = new Date(y, m, 0).getDate();
      const pad = (n: number) => String(n).padStart(2, '0');
      return `${y}-${pad(m)}-01 to ${y}-${pad(m)}-${pad(lastDay)}`;
    } catch {
      return '';
    }
  };

  const ocsPeriod = getOcsPeriod(evaluationMonth);

  return (
    <header className="bg-slate-900 text-white border-b border-slate-800 shadow-md">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-4">
        {/* Left: Branding */}
        <div className="flex items-center gap-3.5">
          <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-blue-600 to-indigo-700 flex items-center justify-center shadow-inner border border-blue-400/30">
            <Layers className="w-5 h-5 text-white" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-lg font-bold tracking-tight text-white leading-tight">
                EIL Manhour Automation & Audit
              </h1>
              <span className="px-2 py-0.5 text-[11px] font-semibold bg-blue-950 text-blue-300 rounded border border-blue-700/50">
                EIL Rules Engine
              </span>
            </div>
            <p className="text-xs text-slate-400">
              Deterministic 6-Step Manhour Reconciliation, Audit & Excel Automation
            </p>
          </div>
        </div>

        {/* Center/Right: Evaluation Month & Project Controls */}
        <div className="flex flex-wrap items-center gap-3">
          {/* Evaluation Month Input */}
          <div className="flex items-center gap-2 bg-slate-800/90 border border-slate-700/80 rounded-lg px-3 py-1.5 shadow-sm">
            <Calendar className="w-4 h-4 text-blue-400" />
            <div className="flex flex-col">
              <span className="text-[10px] font-medium text-slate-400 uppercase tracking-wider">
                Evaluation Month
              </span>
              <input
                type="month"
                id="evaluation-month-picker"
                value={evaluationMonth}
                onChange={(e) => onEvaluationMonthChange(e.target.value)}
                className="bg-transparent text-sm font-semibold text-slate-100 focus:outline-none cursor-pointer"
              />
            </div>
            {ocsPeriod && (
              <div className="hidden lg:block border-l border-slate-700 pl-2.5 ml-1">
                <span className="text-[9px] text-slate-400 block font-medium">Active Cycle Window:</span>
                <span className="text-[11px] font-mono text-emerald-400">{ocsPeriod}</span>
              </div>
            )}
          </div>

          {/* Project Management Actions */}
          <div className="flex items-center gap-1.5">
            <button
              id="save-project-btn"
              onClick={onOpenSaveModal}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-md border border-slate-700 transition"
              title="Save project snapshot with uploaded workbooks and settings"
            >
              <Save className="w-3.5 h-3.5 text-emerald-400" />
              Save Project
            </button>

            <button
              id="open-project-btn"
              onClick={onOpenLoadModal}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-md border border-slate-700 transition"
              title="Open a previously saved project"
            >
              <FolderOpen className="w-3.5 h-3.5 text-amber-400" />
              Open Project
            </button>

            {hasWorkbooks && (
              <button
                id="reset-workbooks-btn"
                onClick={onReset}
                className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-md transition"
                title="Reset active session workbooks"
              >
                <RotateCcw className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      </div>
    </header>
  );
};
