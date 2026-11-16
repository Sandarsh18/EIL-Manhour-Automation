import React, { useState } from 'react';
import { X, Save, FolderOpen, Trash2, Calendar, Check, AlertCircle } from 'lucide-react';
import { SavedProject } from '../types/eil.ts';

interface SaveModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (name: string, description: string) => Promise<void>;
}

export const SaveProjectModal: React.FC<SaveModalProps> = ({ isOpen, onClose, onSave }) => {
  const [name, setName] = useState('');
  const [description, setDescription] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) {
      setError('Please provide a project name.');
      return;
    }
    try {
      setIsSaving(true);
      setError(null);
      await onSave(name.trim(), description.trim());
      setName('');
      setDescription('');
      onClose();
    } catch (err: any) {
      setError(err.message || 'Failed to save project.');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-md overflow-hidden">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2 text-slate-800 font-bold text-sm">
            <Save className="w-4 h-4 text-emerald-600" />
            Save EIL Project Snapshot
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700">
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          {error && (
            <div className="p-2.5 bg-rose-50 border border-rose-200 rounded text-xs text-rose-700 flex items-center gap-1.5">
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
              {error}
            </div>
          )}

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              Project Name <span className="text-rose-500">*</span>
            </label>
            <input
              type="text"
              id="project-name-input"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="e.g. IOCL Panipat Refinery Audit - Aug 2026"
              className="w-full text-xs px-3 py-2 border border-slate-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              autoFocus
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              Description / Audit Notes (Optional)
            </label>
            <textarea
              id="project-desc-input"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Notes on evaluation cycle, contractor details, or revisions..."
              rows={3}
              className="w-full text-xs px-3 py-2 border border-slate-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none resize-none"
            />
          </div>

          <div className="pt-2 flex justify-end gap-2 border-t border-slate-100">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100 rounded-md transition"
            >
              Cancel
            </button>
            <button
              type="submit"
              id="confirm-save-project-btn"
              disabled={isSaving}
              className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-md transition flex items-center gap-1.5 disabled:opacity-50 shadow-xs"
            >
              <Check className="w-3.5 h-3.5" />
              {isSaving ? 'Saving...' : 'Save Project'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

interface LoadModalProps {
  isOpen: boolean;
  onClose: () => void;
  projects: SavedProject[];
  onLoad: (id: string) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
}

export const LoadProjectModal: React.FC<LoadModalProps> = ({
  isOpen,
  onClose,
  projects,
  onLoad,
  onDelete,
}) => {
  const [loadingId, setLoadingId] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleLoad = async (id: string) => {
    try {
      setLoadingId(id);
      await onLoad(id);
      onClose();
    } finally {
      setLoadingId(null);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-2xl max-h-[80vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2 text-slate-800 font-bold text-sm">
            <FolderOpen className="w-4 h-4 text-amber-500" />
            Open Saved EIL Project
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700">
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-5 flex-1 overflow-y-auto">
          {projects.length === 0 ? (
            <div className="text-center py-12 text-slate-400 text-xs">
              No saved projects found. Save your current session to resume later.
            </div>
          ) : (
            <div className="space-y-2.5">
              {projects.map((proj) => {
                const uploadedCount = [
                  proj.workbooksState?.excel1,
                  proj.workbooksState?.excel2,
                  proj.workbooksState?.excel3,
                ].filter(Boolean).length;

                return (
                  <div
                    key={proj.id}
                    className="p-3.5 border border-slate-200 rounded-lg hover:border-blue-300 hover:bg-slate-50/60 transition flex items-center justify-between gap-4"
                  >
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2">
                        <h4 className="text-xs font-bold text-slate-900 truncate">{proj.name}</h4>
                        <span className="px-1.5 py-0.2 text-[10px] bg-slate-100 text-slate-600 rounded border border-slate-200">
                          Step {proj.currentStep || 1}
                        </span>
                        <span className="px-1.5 py-0.2 text-[10px] bg-blue-50 text-blue-700 rounded border border-blue-200">
                          {uploadedCount}/3 workbooks
                        </span>
                      </div>
                      {proj.description && (
                        <p className="text-[11px] text-slate-500 mt-1 line-clamp-1">{proj.description}</p>
                      )}
                      <div className="text-[10px] text-slate-400 mt-1.5 flex items-center gap-3">
                        <span className="flex items-center gap-1">
                          <Calendar className="w-3 h-3 text-slate-400" />
                          Eval Month: {proj.evaluationMonth || '2026-08'}
                        </span>
                        <span>•</span>
                        <span>Saved: {new Date(proj.updatedAt).toLocaleString()}</span>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        onClick={() => handleLoad(proj.id)}
                        disabled={loadingId === proj.id}
                        className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-md transition shadow-2xs"
                      >
                        {loadingId === proj.id ? 'Loading...' : 'Restore'}
                      </button>
                      <button
                        onClick={() => onDelete(proj.id)}
                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition"
                        title="Delete project"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md transition"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};
