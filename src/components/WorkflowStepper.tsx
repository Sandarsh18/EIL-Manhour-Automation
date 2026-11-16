import React from 'react';
import { Upload, Table, CheckSquare, Calculator, CheckCircle, Download, ChevronRight } from 'lucide-react';

interface WorkflowStepperProps {
  currentStep: number;
  maxUnlockedStep?: number;
  onStepClick?: (step: number) => void;
  canProceedToStep2: boolean;
}

export const WorkflowStepper: React.FC<WorkflowStepperProps> = ({
  currentStep,
  maxUnlockedStep = 1,
  onStepClick,
  canProceedToStep2,
}) => {
  const steps = [
    { number: 1, title: 'Upload Workbooks', subtitle: '3 Source Files', icon: Upload },
    { number: 2, title: 'Inspect & Map', subtitle: 'Dynamic Columns', icon: Table },
    { number: 3, title: 'Select Jobs', subtitle: 'Reconciliation', icon: CheckSquare },
    { number: 4, title: 'Rule Analysis', subtitle: 'EIL Calculations', icon: Calculator },
    { number: 5, title: 'Review & Approve', subtitle: 'Auditor Sign-off', icon: CheckCircle },
    { number: 6, title: 'Output Excel', subtitle: 'Template Export', icon: Download },
  ];

  return (
    <div className="bg-white border-b border-slate-200 py-3 shadow-xs">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <nav aria-label="Workflow progress" className="overflow-x-auto">
          <ol className="flex items-center justify-between min-w-[720px] gap-2">
            {steps.map((step, idx) => {
              const Icon = step.icon;
              const isActive = step.number === currentStep;
              const isCompleted = step.number < currentStep;
              const isAccessible = step.number <= maxUnlockedStep;

              return (
                <li key={step.number} className="flex-1 flex items-center">
                  <button
                    type="button"
                    disabled={!isAccessible}
                    onClick={() => isAccessible && onStepClick?.(step.number)}
                    className={`flex items-center gap-2.5 px-3 py-2 rounded-lg w-full transition text-left select-none ${
                      isActive
                        ? 'bg-blue-50/80 border border-blue-300/80 text-blue-900 ring-1 ring-blue-200'
                        : isCompleted
                        ? 'bg-slate-50 border border-slate-200 text-slate-700 hover:bg-slate-100 cursor-pointer'
                        : isAccessible
                        ? 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 cursor-pointer'
                        : 'text-slate-300 border border-transparent cursor-not-allowed opacity-60'
                    }`}
                  >
                    <div
                      className={`w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold shrink-0 transition ${
                        isActive
                          ? 'bg-blue-600 text-white shadow-xs ring-2 ring-blue-100'
                          : isCompleted
                          ? 'bg-emerald-600 text-white'
                          : isAccessible
                          ? 'bg-slate-200 text-slate-700'
                          : 'bg-slate-100 text-slate-400'
                      }`}
                    >
                      {isCompleted ? '✓' : <Icon className="w-3.5 h-3.5" />}
                    </div>

                    <div className="min-w-0">
                      <div className="text-xs font-semibold truncate leading-tight flex items-center gap-1">
                        <span>Step {step.number}</span>
                        {isActive && (
                          <span className="inline-block w-1.5 h-1.5 rounded-full bg-blue-600 animate-pulse"></span>
                        )}
                      </div>
                      <div className="text-[11px] truncate text-slate-500">{step.title}</div>
                    </div>
                  </button>

                  {idx < steps.length - 1 && (
                    <ChevronRight className="w-4 h-4 text-slate-300 mx-1 shrink-0" />
                  )}
                </li>
              );
            })}
          </ol>
        </nav>
      </div>
    </div>
  );
};
