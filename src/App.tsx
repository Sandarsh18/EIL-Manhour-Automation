import React, { useEffect, useState } from 'react';
import {
  WorkbooksState,
  WorkbookType,
  SavedProject,
  MappingsState,
  JobReconciliationSummary,
  JobCalculationResult,
  ReviewItem,
  OutputValidationReport,
} from './types/eil.ts';
import { EilApi } from './services/api.ts';
import { Header } from './components/Header.tsx';
import { WorkflowStepper } from './components/WorkflowStepper.tsx';
import { Step1UploadView } from './components/Step1UploadView.tsx';
import { Step2MappingView } from './components/Step2MappingView.tsx';
import { Step3JobSelectView } from './components/Step3JobSelectView.tsx';
import { Step4RuleAnalysisView } from './components/Step4RuleAnalysisView.tsx';
import { Step5ReviewApproveView } from './components/Step5ReviewApproveView.tsx';
import { Step6OutputView } from './components/Step6OutputView.tsx';
import { SaveProjectModal, LoadProjectModal } from './components/ProjectModal.tsx';
import { AlertCircle, CheckCircle2 } from 'lucide-react';

export default function App() {
  const [currentStep, setCurrentStep] = useState<number>(1);
  const [maxUnlockedStep, setMaxUnlockedStep] = useState<number>(1);
  const [evaluationMonth, setEvaluationMonth] = useState<string>('2026-08');

  const [workbooks, setWorkbooks] = useState<WorkbooksState>({
    excel1: null,
    excel2: null,
    excel3: null,
  });

  const [mappings, setMappings] = useState<MappingsState>({
    excel1: {
      workbookType: 'excel1',
      selectedSheet: '',
      fields: [],
      isValid: false,
      validationErrors: [],
      validationWarnings: [],
    },
    excel2: {
      workbookType: 'excel2',
      selectedSheet: '',
      fields: [],
      isValid: false,
      validationErrors: [],
      validationWarnings: [],
    },
    excel3: {
      workbookType: 'excel3',
      selectedSheet: '',
      fields: [],
      isValid: false,
      validationErrors: [],
      validationWarnings: [],
    },
  });

  const [reconciliation, setReconciliation] = useState<JobReconciliationSummary | null>(null);
  const [selectedJobs, setSelectedJobs] = useState<string[]>([]);
  const [calculations, setCalculations] = useState<Record<string, JobCalculationResult>>({});
  const [reviewItems, setReviewItems] = useState<Record<string, ReviewItem>>({});
  const [outputReport, setOutputReport] = useState<OutputValidationReport | null>(null);

  const [projects, setProjects] = useState<SavedProject[]>([]);
  const [isSaveModalOpen, setIsSaveModalOpen] = useState(false);
  const [isLoadModalOpen, setIsLoadModalOpen] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const [toastMessage, setToastMessage] = useState<{
    type: 'success' | 'error' | 'info';
    text: string;
  } | null>(null);

  // Start with empty workbook slots on app load while preserving Saved Projects
  useEffect(() => {
    loadInitialData();
  }, []);

  const loadInitialData = async () => {
    try {
      await EilApi.resetWorkbooks();
      setWorkbooks({ excel1: null, excel2: null, excel3: null });
      setCalculations({});
      setReviewItems({});
      setOutputReport(null);
      setReconciliation(null);
      setSelectedJobs([]);
      setCurrentStep(1);
      setMaxUnlockedStep(1);

      const session = await EilApi.getSessionState();
      if (session) {
        if (session.mappings) setMappings(session.mappings);
        if (session.evaluationMonth) setEvaluationMonth(session.evaluationMonth);
      }

      const projList = await EilApi.listProjects();
      if (projList) setProjects(projList);
    } catch (err: any) {
      console.warn('Could not load initial server state:', err);
    }
  };

  const showToast = (type: 'success' | 'error' | 'info', text: string) => {
    setToastMessage({ type, text });
    setTimeout(() => {
      setToastMessage(null);
    }, 4000);
  };

  const unlockStep = (step: number) => {
    setMaxUnlockedStep((prev) => Math.max(prev, step));
    setCurrentStep(step);
    EilApi.syncWorkflowStep(step, evaluationMonth).catch(console.error);
  };

  // Step 1: Upload / Replace / Delete
  const handleUpload = async (type: WorkbookType, file: File) => {
    try {
      const meta = await EilApi.uploadWorkbook(type, file);
      const nextWorkbooks = { ...workbooks, [type]: meta };
      setWorkbooks(nextWorkbooks);

      // Fetch refreshed mapping configs
      const updatedMappings = await EilApi.getMappings();
      setMappings(updatedMappings);

      // Invalidate downstream calculations, reconciliation, and job selection
      setCalculations({});
      setReviewItems({});
      setOutputReport(null);
      setReconciliation(null);
      setSelectedJobs([]);

      const allReady =
        Boolean(nextWorkbooks.excel1 && nextWorkbooks.excel2 && nextWorkbooks.excel3) &&
        [nextWorkbooks.excel1, nextWorkbooks.excel2, nextWorkbooks.excel3].every(
          (w) => w?.validationStatus !== 'ERROR'
        );
      setMaxUnlockedStep(allReady ? 2 : 1);

      showToast('success', `Uploaded & parsed ${meta.fileName} successfully.`);
    } catch (err: any) {
      showToast('error', err.message || `Failed to upload ${type}`);
      throw err;
    }
  };

  const handleDelete = async (type: WorkbookType) => {
    try {
      await EilApi.deleteWorkbook(type);
      setWorkbooks((prev) => ({ ...prev, [type]: null }));

      const updatedMappings = await EilApi.getMappings();
      setMappings(updatedMappings);

      setCalculations({});
      setReviewItems({});
      setOutputReport(null);
      setReconciliation(null);
      setSelectedJobs([]);
      setCurrentStep(1);
      setMaxUnlockedStep(1);

      showToast('info', `Removed ${type.toUpperCase()} from active session.`);
    } catch (err: any) {
      showToast('error', err.message || `Failed to remove ${type}`);
    }
  };

  const handleReset = async () => {
    try {
      await EilApi.resetWorkbooks();
      setWorkbooks({ excel1: null, excel2: null, excel3: null });
      const updatedMappings = await EilApi.getMappings();
      setMappings(updatedMappings);
      setCalculations({});
      setReviewItems({});
      setOutputReport(null);
      setReconciliation(null);
      setSelectedJobs([]);
      setCurrentStep(1);
      setMaxUnlockedStep(1);
      showToast('info', 'All active workbooks have been reset.');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to reset workbooks.');
    }
  };

  // Step 2 -> Step 3 Transition
  const handleProceedToStep3 = async () => {
    try {
      setIsProcessing(true);
      const summary = await EilApi.reconcileJobs();
      setReconciliation(summary);

      // Default selected jobs to all jobs if none selected
      if (selectedJobs.length === 0) {
        setSelectedJobs(summary.jobs.map((j) => j.jobNumber));
      }

      unlockStep(3);
      showToast('success', `Reconciled ${summary.totalJobs} jobs across workbooks.`);
    } catch (err: any) {
      showToast('error', err.message || 'Failed to reconcile jobs');
    } finally {
      setIsProcessing(false);
    }
  };

  // Step 3 -> Step 4 Transition
  const handleProceedToStep4 = async () => {
    try {
      setIsProcessing(true);
      const results = await EilApi.runRuleAnalysis(selectedJobs, evaluationMonth);
      setCalculations(results);
      unlockStep(4);
      showToast('success', `Evaluated rules for ${Object.keys(results).length} jobs.`);
    } catch (err: any) {
      showToast('error', err.message || 'Rule analysis execution failed');
    } finally {
      setIsProcessing(false);
    }
  };

  // Re-run rule analysis in Step 4
  const handleRunAnalysis = async () => {
    try {
      setIsProcessing(true);
      const results = await EilApi.runRuleAnalysis(selectedJobs, evaluationMonth);
      setCalculations(results);
      showToast('success', `Rules re-evaluated successfully for ${evaluationMonth}.`);
    } catch (err: any) {
      showToast('error', err.message || 'Failed to re-run analysis');
    } finally {
      setIsProcessing(false);
    }
  };

  // Step 4 -> Step 5 Transition
  const handleProceedToStep5 = async () => {
    try {
      setIsProcessing(true);
      const items = await EilApi.getReviewItems();
      setReviewItems(items);
      unlockStep(5);
    } catch (err: any) {
      showToast('error', err.message || 'Failed to load review items');
    } finally {
      setIsProcessing(false);
    }
  };

  // Step 5 Review Actions
  const handleUpdateReviewItem = async (jobNumber: string, updates: Partial<ReviewItem>) => {
    try {
      const updated = await EilApi.updateReviewItem(jobNumber, updates);
      setReviewItems((prev) => ({ ...prev, [jobNumber]: updated }));
      showToast('info', `Updated review for Job ${jobNumber}`);
    } catch (err: any) {
      showToast('error', err.message || 'Failed to update review item');
    }
  };

  const handleApproveAll = async () => {
    try {
      const updated = await EilApi.approveAllReviews();
      setReviewItems(updated);
      showToast('success', 'All non-blocked jobs approved successfully.');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to approve reviews');
    }
  };

  const handleUnapproveAll = async () => {
    try {
      const updated = await EilApi.unapproveAllReviews();
      setReviewItems(updated);
      showToast('info', 'All jobs reset to Pending approval.');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to reset reviews to Pending');
    }
  };

  const handleDeleteAllReviews = async () => {
    try {
      const updated = await EilApi.deleteAllReviews();
      setReviewItems(updated);
      showToast('info', 'All review jobs cleared from active scope.');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to delete reviews');
    }
  };

  const handleDeleteReviewItem = async (jobNumber: string) => {
    if (!jobNumber || typeof jobNumber !== 'string' || !jobNumber.trim()) return;
    const targetJob = jobNumber.trim();
    try {
      const updated = await EilApi.deleteReviewItem(targetJob);
      setReviewItems(updated);
      setCalculations((prev) => {
        const copy = { ...prev };
        delete copy[targetJob];
        delete copy[targetJob.toUpperCase()];
        return copy;
      });
      setSelectedJobs((prev) => prev.filter((j) => j && j.trim() !== targetJob && j.trim().toUpperCase() !== targetJob.toUpperCase()));
      setOutputReport(null);
      showToast('info', `Job ${targetJob} deleted from authoritative dataset.`);
    } catch (err: any) {
      showToast('error', err.message || 'Failed to delete job review item');
    }
  };

  // Step 5 -> Step 6 Transition & Output Generation
  const handleProceedToStep6 = async () => {
    try {
      setIsProcessing(true);
      const report = await EilApi.generateOutput();
      setOutputReport(report);
      unlockStep(6);
      showToast('success', 'Master output workbook generated successfully!');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to generate output workbook');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleGenerateOutput = async () => {
    try {
      setIsProcessing(true);
      const report = await EilApi.generateOutput();
      setOutputReport(report);
      showToast('success', 'Output workbook regenerated successfully!');
    } catch (err: any) {
      showToast('error', err.message || 'Failed to generate output workbook');
    } finally {
      setIsProcessing(false);
    }
  };

  // Project Persistence
  const handleSaveProject = async (name: string, description: string) => {
    const saved = await EilApi.saveProject(name, description);
    const updatedProjects = await EilApi.listProjects();
    setProjects(updatedProjects);
    showToast('success', `Project "${saved.name}" saved successfully.`);
  };

  const handleLoadProject = async (id: string) => {
    const loaded = await EilApi.loadProject(id);
    if (loaded) {
      setWorkbooks(loaded.workbooksState);
      if (loaded.mappingsState) setMappings(loaded.mappingsState);
      if (loaded.evaluationMonth) setEvaluationMonth(loaded.evaluationMonth);
      if (loaded.selectedJobs) setSelectedJobs(loaded.selectedJobs);
      if (loaded.calculations) setCalculations(loaded.calculations);
      if (loaded.reviewItems) setReviewItems(loaded.reviewItems);
      if (loaded.outputReport) setOutputReport(loaded.outputReport);
      if (loaded.currentStep) {
        setCurrentStep(loaded.currentStep);
        setMaxUnlockedStep(Math.max(maxUnlockedStep, loaded.currentStep));
      }
      showToast('success', `Restored project "${loaded.name}".`);
    }
  };

  const handleDeleteProject = async (id: string) => {
    await EilApi.deleteProject(id);
    const updatedProjects = await EilApi.listProjects();
    setProjects(updatedProjects);
    showToast('info', 'Project deleted.');
  };

  const hasWorkbooks = Boolean(workbooks.excel1 || workbooks.excel2 || workbooks.excel3);
  const canProceedToStep2 =
    Boolean(workbooks.excel1 && workbooks.excel2 && workbooks.excel3) &&
    [workbooks.excel1, workbooks.excel2, workbooks.excel3].every((w) => w?.validationStatus !== 'ERROR');

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 flex flex-col font-sans">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-5 right-5 z-50 animate-in fade-in slide-in-from-bottom-5">
          <div
            className={`flex items-center gap-2.5 px-4 py-2.5 rounded-lg shadow-xl text-xs font-medium border ${
              toastMessage.type === 'success'
                ? 'bg-emerald-900 text-emerald-100 border-emerald-700'
                : toastMessage.type === 'error'
                ? 'bg-rose-900 text-rose-100 border-rose-700'
                : 'bg-slate-900 text-slate-100 border-slate-700'
            }`}
          >
            {toastMessage.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />}
            {toastMessage.type === 'error' && <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />}
            <span>{toastMessage.text}</span>
          </div>
        </div>
      )}

      {/* Top Application Header */}
      <Header
        evaluationMonth={evaluationMonth}
        onEvaluationMonthChange={(month) => {
          setEvaluationMonth(month);
          EilApi.syncWorkflowStep(currentStep, month).catch(console.error);
        }}
        onOpenSaveModal={() => setIsSaveModalOpen(true)}
        onOpenLoadModal={() => setIsLoadModalOpen(true)}
        onReset={handleReset}
        hasWorkbooks={hasWorkbooks}
      />

      {/* Workflow Stepper Bar */}
      <WorkflowStepper
        currentStep={currentStep}
        maxUnlockedStep={maxUnlockedStep}
        onStepClick={(step) => setCurrentStep(step)}
        canProceedToStep2={canProceedToStep2}
      />

      {/* Main View Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {currentStep === 1 && (
          <Step1UploadView
            workbooks={workbooks}
            onUpload={handleUpload}
            onDelete={handleDelete}
            onProceedToStep2={() => {
              unlockStep(2);
            }}
          />
        )}

        {currentStep === 2 && (
          <Step2MappingView
            workbooks={workbooks}
            mappings={mappings}
            onMappingsChange={(updated) => setMappings(updated)}
            onBackToStep1={() => setCurrentStep(1)}
            onProceedToStep3={handleProceedToStep3}
          />
        )}

        {currentStep === 3 && (
          <Step3JobSelectView
            reconciliation={reconciliation}
            selectedJobs={selectedJobs}
            onSelectedJobsChange={setSelectedJobs}
            onBackToStep2={() => setCurrentStep(2)}
            onProceedToStep4={handleProceedToStep4}
          />
        )}

        {currentStep === 4 && (
          <Step4RuleAnalysisView
            calculations={calculations}
            selectedJobs={selectedJobs}
            evaluationMonth={evaluationMonth}
            onEvaluationMonthChange={(m) => {
              setEvaluationMonth(m);
              EilApi.syncWorkflowStep(currentStep, m).catch(console.error);
            }}
            onRunAnalysis={handleRunAnalysis}
            onBackToStep3={() => setCurrentStep(3)}
            onProceedToStep5={handleProceedToStep5}
            isLoading={isProcessing}
          />
        )}

        {currentStep === 5 && (
          <Step5ReviewApproveView
            reviewItems={reviewItems}
            onUpdateReviewItem={handleUpdateReviewItem}
            onApproveAll={handleApproveAll}
            onUnapproveAll={handleUnapproveAll}
            onDeleteAll={handleDeleteAllReviews}
            onDeleteReviewItem={handleDeleteReviewItem}
            onBackToStep4={() => setCurrentStep(4)}
            onProceedToStep6={handleProceedToStep6}
            isGenerating={isProcessing}
          />
        )}

        {currentStep === 6 && (
          <Step6OutputView
            outputReport={outputReport}
            reviewItems={reviewItems}
            evaluationMonth={evaluationMonth}
            onGenerateOutput={handleGenerateOutput}
            onBackToStep5={() => setCurrentStep(5)}
            isGenerating={isProcessing}
          />
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 py-3 text-center text-xs text-slate-500">
        <div className="max-w-7xl mx-auto px-4 flex items-center justify-between">
          <span>EIL Manhour Automation & Audit • Engineers India Limited Compliance Engine</span>
          <span className="font-mono text-[11px] text-slate-400">Production Build v2.0</span>
        </div>
      </footer>

      {/* Modals */}
      <SaveProjectModal
        isOpen={isSaveModalOpen}
        onClose={() => setIsSaveModalOpen(false)}
        onSave={handleSaveProject}
      />

      <LoadProjectModal
        isOpen={isLoadModalOpen}
        onClose={() => setIsLoadModalOpen(false)}
        projects={projects}
        onLoad={handleLoadProject}
        onDelete={handleDeleteProject}
      />
    </div>
  );
}
