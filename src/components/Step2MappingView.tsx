import React, { useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  AlertTriangle,
  AlertCircle,
  Sliders,
  FileInput,
  FileOutput,
  RefreshCw,
  Sparkles,
  Info,
  Layers,
  Edit3,
} from 'lucide-react';
import {
  WorkbooksState,
  WorkbookType,
  MappingsState,
  ColumnMappingField,
} from '../types/eil.ts';
import { EilApi } from '../services/api.ts';

interface Step2MappingViewProps {
  workbooks: WorkbooksState;
  mappings: MappingsState;
  onMappingsChange: (mappings: MappingsState) => void;
  onBackToStep1: () => void;
  onProceedToStep3: () => void;
}

type ViewMode = 'excel1' | 'excel2' | 'excel3' | 'all';

export const Step2MappingView: React.FC<Step2MappingViewProps> = ({
  workbooks,
  mappings,
  onMappingsChange,
  onBackToStep1,
  onProceedToStep3,
}) => {
  const [activeTab, setActiveTab] = useState<ViewMode>('excel1');
  const [isValidating, setIsValidating] = useState(false);
  const [validationResult, setValidationResult] = useState<{
    isValid: boolean;
    errors: string[];
    warnings: string[];
  } | null>(null);

  // Handle dropdown selection change for any workbook type
  const handleFieldChange = async (
    wbType: WorkbookType,
    canonicalField: string,
    newMappedHeader: string
  ) => {
    const currentMapping = mappings[wbType];
    if (!currentMapping) return;

    const updatedFields = currentMapping.fields.map((f) => {
      if (f.canonicalField === canonicalField) {
        return {
          ...f,
          mappedHeader: newMappedHeader === '' ? null : newMappedHeader,
          confidence: 'MANUAL' as const,
          matchReason: 'Manually assigned by user',
        };
      }
      return f;
    });

    try {
      const updated = await EilApi.updateMapping(
        wbType,
        currentMapping.selectedSheet,
        updatedFields
      );
      onMappingsChange(updated);
      setValidationResult(null); // Clear previous validation state to force re-validation
    } catch (err: any) {
      console.error('Failed to update mapping:', err);
    }
  };

  const handleSheetChange = async (wbType: WorkbookType, newSheetName: string) => {
    const currentWb = workbooks[wbType];
    const currentMapping = mappings[wbType];
    if (!currentWb || !currentMapping) return;

    const sheetObj = currentWb.sheets.find((s) => s.name === newSheetName);
    if (!sheetObj) return;

    try {
      const updated = await EilApi.updateMapping(
        wbType,
        newSheetName,
        currentMapping.fields || []
      );
      onMappingsChange(updated);
      setValidationResult(null);
    } catch (err: any) {
      console.error('Failed to switch worksheet:', err);
    }
  };

  const handleValidateMapping = async () => {
    try {
      setIsValidating(true);
      const res = await EilApi.validateMappings();
      setValidationResult(res);
    } catch (err: any) {
      setValidationResult({
        isValid: false,
        errors: [err.message || 'Validation request failed'],
        warnings: [],
      });
    } finally {
      setIsValidating(false);
    }
  };

  const getTabStatus = (type: WorkbookType) => {
    const m = mappings[type];
    if (!m) return { text: 'Missing', color: 'text-slate-400', bg: 'bg-slate-100 text-slate-600' };
    if (m.isValid) return { text: 'Valid', color: 'text-emerald-600', bg: 'bg-emerald-100 text-emerald-800' };
    return { text: 'Needs Mapping', color: 'text-amber-600', bg: 'bg-amber-100 text-amber-800' };
  };

  // Check unique targets for Excel 3
  const getExcel3DuplicateHeaders = (): Set<string> => {
    const dupes = new Set<string>();
    const counts = new Map<string, number>();
    const fields = mappings.excel3?.fields || [];
    for (const f of fields) {
      if (f.mappedHeader && !f.mappedHeader.startsWith('[')) {
        const c = (counts.get(f.mappedHeader) || 0) + 1;
        counts.set(f.mappedHeader, c);
        if (c > 1) dupes.add(f.mappedHeader);
      }
    }
    return dupes;
  };

  const excel3Dupes = getExcel3DuplicateHeaders();

  const allWorkbooksValid =
    Boolean(mappings.excel1?.isValid) &&
    Boolean(mappings.excel2?.isValid) &&
    Boolean(mappings.excel3?.isValid) &&
    excel3Dupes.size === 0;

  const canProceed = allWorkbooksValid && (!validationResult || validationResult.isValid);

  // Helper badge styling for Excel 3 logic origins
  const getOriginBadge = (field: ColumnMappingField) => {
    if (field.calculationSource === 'CALCULATED FROM EXCEL 1') {
      return (
        <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-blue-100 text-blue-800 border border-blue-200">
          CALCULATED FROM EXCEL 1
        </span>
      );
    }
    if (field.calculationSource === 'CALCULATED FROM EXCEL 2') {
      return (
        <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-indigo-100 text-indigo-800 border border-indigo-200">
          CALCULATED FROM EXCEL 2
        </span>
      );
    }
    if (field.calculationSource === 'CALCULATED') {
      return (
        <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-amber-100 text-amber-800 border border-amber-200">
          CALCULATED
        </span>
      );
    }
    if (field.calculationSource === 'AUDIT NOTES') {
      return (
        <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
          AUDIT NOTES
        </span>
      );
    }
    return (
      <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
        TARGET
      </span>
    );
  };

  // Render Section A: Excel 1 Source Mapping
  const renderSectionA = () => {
    const wb = workbooks.excel1;
    const mapping = mappings.excel1;
    const sheet = wb?.sheets.find((s) => s.name === mapping?.selectedSheet) || wb?.sheets[0];

    return (
      <div id="section-a-excel1" className="bg-white rounded-xl border border-slate-200 p-6 space-y-5 shadow-xs">
        {/* Section Header */}
        <div className="flex flex-wrap items-start justify-between gap-4 pb-4 border-b border-slate-200">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2 py-0.5 rounded text-[11px] font-bold uppercase bg-blue-100 text-blue-800 border border-blue-200">
                SECTION A
              </span>
              <h3 className="text-base font-bold text-slate-900">
                EXCEL 1 — SOURCE COLUMN MAPPING
              </h3>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Consolidated Report • Map exactly the 5 source input columns. Do not map calculated fields here.
            </p>
          </div>

          {/* Worksheet selector */}
          {wb && mapping && (
            <div className="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200">
              <label className="text-xs font-semibold text-slate-600">Worksheet:</label>
              <select
                id="select-sheet-excel1"
                value={mapping.selectedSheet}
                onChange={(e) => handleSheetChange('excel1', e.target.value)}
                className="text-xs bg-white border border-slate-300 rounded px-2.5 py-1 text-slate-800 font-medium focus:ring-2 focus:ring-blue-500"
              >
                {wb.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name} ({s.rowCount} rows)
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Source Inputs Table */}
        {wb && mapping ? (
          <div className="space-y-4">
            <div className="border border-slate-200 rounded-lg overflow-hidden divide-y divide-slate-200">
              <div className="bg-slate-100/80 px-4 py-2.5 text-xs font-bold text-slate-700 grid grid-cols-12 gap-4">
                <div className="col-span-5">Canonical Source Input Field</div>
                <div className="col-span-4">Mapped Source Column (Worksheet Header)</div>
                <div className="col-span-3">Detection Status</div>
              </div>

              {mapping.fields.map((field) => (
                <div
                  key={field.canonicalField}
                  className="px-4 py-3 bg-white hover:bg-slate-50/60 grid grid-cols-12 gap-4 items-center transition"
                >
                  {/* Field Label & info */}
                  <div className="col-span-5">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-slate-900">{field.label}</span>
                      <span className="text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-200 px-1.5 py-0.2 rounded">
                        REQUIRED
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-500 mt-0.5 leading-tight">
                      {field.description}
                    </p>
                  </div>

                  {/* Dropdown */}
                  <div className="col-span-4">
                    <select
                      id={`map-select-excel1-${field.canonicalField}`}
                      value={field.mappedHeader || ''}
                      onChange={(e) => handleFieldChange('excel1', field.canonicalField, e.target.value)}
                      className={`w-full text-xs rounded-md px-3 py-1.5 border transition font-mono ${
                        !field.mappedHeader
                          ? 'border-rose-300 bg-rose-50/40 text-rose-900 font-medium'
                          : 'border-slate-300 bg-white text-slate-900 focus:ring-blue-500'
                      }`}
                    >
                      <option value="">-- Not Mapped (Select Source Column) --</option>
                      {sheet?.headers.map((h, i) => (
                        <option key={i} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Confidence Badge */}
                  <div className="col-span-3">
                    {field.mappedHeader ? (
                      <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                        <CheckCircle2 className="w-3 h-3" />
                        {field.confidence === 'MANUAL' ? 'User Selected' : 'Auto-detected'}
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                        <AlertCircle className="w-3 h-3" />
                        Unmapped
                      </span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="text-center py-8 text-xs text-slate-400">
            Excel 1 (Consolidated Report) is not uploaded. Please return to Step 1.
          </div>
        )}
      </div>
    );
  };

  // Render Section B: Excel 2 Source Mapping
  const renderSectionB = () => {
    const wb = workbooks.excel2;
    const mapping = mappings.excel2;
    const sheet = wb?.sheets.find((s) => s.name === mapping?.selectedSheet) || wb?.sheets[0];

    return (
      <div id="section-b-excel2" className="bg-white rounded-xl border border-slate-200 p-6 space-y-5 shadow-xs">
        {/* Section Header */}
        <div className="flex flex-wrap items-start justify-between gap-4 pb-4 border-b border-slate-200">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2 py-0.5 rounded text-[11px] font-bold uppercase bg-indigo-100 text-indigo-800 border border-indigo-200">
                SECTION B
              </span>
              <h3 className="text-base font-bold text-slate-900">
                EXCEL 2 — SOURCE COLUMN MAPPING
              </h3>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Inspection Call Log • Map exactly the 6 source input columns. Shared "No. of Days" feeds both Inspection and QAP/Others.
            </p>
          </div>

          {/* Worksheet selector */}
          {wb && mapping && (
            <div className="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200">
              <label className="text-xs font-semibold text-slate-600">Worksheet:</label>
              <select
                id="select-sheet-excel2"
                value={mapping.selectedSheet}
                onChange={(e) => handleSheetChange('excel2', e.target.value)}
                className="text-xs bg-white border border-slate-300 rounded px-2.5 py-1 text-slate-800 font-medium focus:ring-2 focus:ring-blue-500"
              >
                {wb.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name} ({s.rowCount} rows)
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Source Inputs Table */}
        {wb && mapping ? (
          <div className="space-y-4">
            <div className="border border-slate-200 rounded-lg overflow-hidden divide-y divide-slate-200">
              <div className="bg-slate-100/80 px-4 py-2.5 text-xs font-bold text-slate-700 grid grid-cols-12 gap-4">
                <div className="col-span-5">Canonical Source Input Field</div>
                <div className="col-span-4">Mapped Source Column (Worksheet Header)</div>
                <div className="col-span-3">Detection Status</div>
              </div>

              {mapping.fields.map((field) => (
                <div
                  key={field.canonicalField}
                  className="px-4 py-3 bg-white hover:bg-slate-50/60 grid grid-cols-12 gap-4 items-center transition"
                >
                  {/* Field Label & info */}
                  <div className="col-span-5">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-slate-900">{field.label}</span>
                      <span className="text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-200 px-1.5 py-0.2 rounded">
                        REQUIRED
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-500 mt-0.5 leading-tight">
                      {field.description}
                    </p>
                  </div>

                  {/* Dropdown */}
                  <div className="col-span-4">
                    <select
                      id={`map-select-excel2-${field.canonicalField}`}
                      value={field.mappedHeader || ''}
                      onChange={(e) => handleFieldChange('excel2', field.canonicalField, e.target.value)}
                      className={`w-full text-xs rounded-md px-3 py-1.5 border transition font-mono ${
                        !field.mappedHeader
                          ? 'border-rose-300 bg-rose-50/40 text-rose-900 font-medium'
                          : 'border-slate-300 bg-white text-slate-900 focus:ring-blue-500'
                      }`}
                    >
                      <option value="">-- Not Mapped (Select Source Column) --</option>
                      {sheet?.headers.map((h, i) => (
                        <option key={i} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Confidence Badge */}
                  <div className="col-span-3">
                    {field.mappedHeader ? (
                      <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                        <CheckCircle2 className="w-3 h-3" />
                        {field.confidence === 'MANUAL' ? 'User Selected' : 'Auto-detected'}
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                        <AlertCircle className="w-3 h-3" />
                        Unmapped
                      </span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="text-center py-8 text-xs text-slate-400">
            Excel 2 (Inspection Call Log) is not uploaded. Please return to Step 1.
          </div>
        )}
      </div>
    );
  };

  // Render Section C: Excel 3 Output Target Mapping
  const renderSectionC = () => {
    const wb = workbooks.excel3;
    const mapping = mappings.excel3;
    const sheet = wb?.sheets.find((s) => s.name === mapping?.selectedSheet) || wb?.sheets[0];

    return (
      <div id="section-c-excel3" className="bg-white rounded-xl border border-slate-200 p-6 space-y-6 shadow-xs">
        {/* Section Header */}
        <div className="flex flex-wrap items-start justify-between gap-4 pb-4 border-b border-slate-200">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2 py-0.5 rounded text-[11px] font-bold uppercase bg-emerald-100 text-emerald-800 border border-emerald-200">
                SECTION C
              </span>
              <h3 className="text-base font-bold text-slate-900">
                EXCEL 3 — OUTPUT TARGET MAPPING
              </h3>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Master Template • Specify where each calculated value is written in Excel 3. Each output target must map to a unique column.
            </p>
          </div>

          {/* Worksheet selector */}
          {wb && mapping && (
            <div className="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200">
              <label className="text-xs font-semibold text-slate-600">Worksheet:</label>
              <select
                id="select-sheet-excel3"
                value={mapping.selectedSheet}
                onChange={(e) => handleSheetChange('excel3', e.target.value)}
                className="text-xs bg-white border border-slate-300 rounded px-2.5 py-1 text-slate-800 font-medium focus:ring-2 focus:ring-emerald-500"
              >
                {wb.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name} ({s.rowCount} rows)
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Output Targets Table */}
        {wb && mapping ? (
          <div className="space-y-6">
            <div className="border border-slate-200 rounded-lg overflow-hidden divide-y divide-slate-200">
              {/* Header labeled strictly 'Output Target Field' per requirement 17 */}
              <div className="bg-slate-100/80 px-4 py-2.5 text-xs font-bold text-slate-700 grid grid-cols-12 gap-4">
                <div className="col-span-5">Output Target Field</div>
                <div className="col-span-4">Target Excel 3 Column</div>
                <div className="col-span-3">Mapping Status</div>
              </div>

              {mapping.fields.map((field) => {
                const isDupe = field.mappedHeader && excel3Dupes.has(field.mappedHeader);

                return (
                  <div
                    key={field.canonicalField}
                    className="px-4 py-3 bg-white hover:bg-slate-50/60 grid grid-cols-12 gap-4 items-center transition"
                  >
                    {/* Field Label & Formula / Logic */}
                    <div className="col-span-5 pr-2">
                      <div className="flex flex-wrap items-center gap-2 mb-1">
                        <span className="text-xs font-bold text-slate-900">{field.label}</span>
                        {getOriginBadge(field)}
                        {field.required ? (
                          <span className="text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-200 px-1.5 py-0.2 rounded">
                            REQUIRED
                          </span>
                        ) : (
                          <span className="text-[10px] font-medium text-slate-500 bg-slate-100 border border-slate-200 px-1.5 py-0.2 rounded">
                            OPTIONAL
                          </span>
                        )}
                      </div>

                      {/* Display explicit calculation logic prominently per requirement 19 & 29 */}
                      <p className="text-[11px] text-slate-600 font-mono leading-tight bg-slate-50 px-2 py-1 rounded border border-slate-200/80 mt-1">
                        {field.description}
                      </p>
                    </div>

                    {/* Column Dropdown */}
                    <div className="col-span-4">
                      <select
                        id={`map-select-excel3-${field.canonicalField}`}
                        value={field.mappedHeader || ''}
                        onChange={(e) => handleFieldChange('excel3', field.canonicalField, e.target.value)}
                        className={`w-full text-xs rounded-md px-3 py-1.5 border transition font-mono ${
                          isDupe
                            ? 'border-rose-400 bg-rose-50 text-rose-900 font-bold focus:ring-rose-500'
                            : !field.mappedHeader && field.required && field.canonicalField !== 'orders_fd' && field.canonicalField !== 'ocs_done'
                            ? 'border-rose-300 bg-rose-50/40 text-rose-900 font-medium'
                            : 'border-slate-300 bg-white text-slate-900 focus:ring-emerald-500'
                        }`}
                      >
                        {(field.canonicalField === 'orders_fd' || field.canonicalField === 'ocs_done') && (
                          <option value="[Auto-create in Excel 3]">✨ [Auto-create column in Excel 3]</option>
                        )}
                        <option value="">-- Select Target Template Column --</option>
                        {sheet?.headers.map((h, i) => (
                          <option key={i} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                      {isDupe && (
                        <p className="text-[10px] text-rose-600 font-semibold mt-1">
                          Duplicate: "{field.mappedHeader}" is mapped multiple times. Each target must be unique.
                        </p>
                      )}
                    </div>

                    {/* Status / Confidence */}
                    <div className="col-span-3">
                      {isDupe ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                          <AlertCircle className="w-3 h-3" />
                          Duplicate Target
                        </span>
                      ) : field.mappedHeader && field.mappedHeader.startsWith('[Auto-create') ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                          <Sparkles className="w-3 h-3 text-emerald-600" />
                          Will be created in template
                        </span>
                      ) : field.mappedHeader ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                          <CheckCircle2 className="w-3 h-3" />
                          {field.confidence === 'MANUAL' ? 'User Selected' : 'Target Identified'}
                        </span>
                      ) : (field.canonicalField === 'orders_fd' || field.canonicalField === 'ocs_done') ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                          <Sparkles className="w-3 h-3 text-emerald-600" />
                          Will be created in template
                        </span>
                      ) : field.required ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                          <AlertCircle className="w-3 h-3" />
                          Target Missing
                        </span>
                      ) : (
                        <span className="text-[11px] text-slate-400">Optional unmapped</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Dedicated MEETING Block per Requirement 13, 21, and 29 */}
            <div
              id="block-meeting-manual-entry"
              className="bg-amber-50/70 border border-amber-200 rounded-xl p-4.5 space-y-2 text-xs"
            >
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <Edit3 className="w-4 h-4 text-amber-700" />
                  <span className="font-bold text-sm text-slate-900">MEETING</span>
                  <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-amber-200/80 text-amber-900 border border-amber-300">
                    MANUAL ENTRY
                  </span>
                </div>
                <span className="text-[11px] font-medium text-amber-800">
                  No source mapping required
                </span>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-slate-700 mt-1">
                <div className="bg-white/80 p-2.5 rounded-lg border border-amber-200/60">
                  <div className="font-semibold text-slate-800 mb-0.5">Workflow & Execution:</div>
                  <p className="text-slate-600 leading-relaxed text-[11px]">
                    The user will enter Meeting values manually after calculation in Step 4/Step 5. Recalculation preserves Meeting values.
                  </p>
                </div>
                <div className="bg-white/80 p-2.5 rounded-lg border border-amber-200/60">
                  <div className="font-semibold text-slate-800 mb-0.5">Output Template Destination:</div>
                  <p className="text-slate-600 leading-relaxed text-[11px]">
                    If the Excel 3 template contains a "Meeting" column, it is populated automatically. If not, the application inserts a "Meeting" column before Total.
                  </p>
                </div>
              </div>

              <div className="text-[11px] text-amber-900/80 font-mono bg-amber-100/60 px-3 py-1.5 rounded">
                Formula contribution: Total = Exp. + Inspn + Others + Meeting
              </div>
            </div>
          </div>
        ) : (
          <div className="text-center py-8 text-xs text-slate-400">
            Excel 3 (Master Template) is not uploaded. Please return to Step 1.
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="space-y-6 pb-24">
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
                Inspect & Map Columns
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Strictly separate Source Input fields (Excel 1 & 2) from Output Target fields (Excel 3 Master Template).
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onBackToStep1}
            className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md transition flex items-center gap-1.5"
          >
            <ArrowLeft className="w-3.5 h-3.5" />
            Back to Step 1
          </button>
        </div>
      </div>

      {/* Section Navigation Tabs Switcher */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        {/* Tab A: Excel 1 */}
        <button
          id="tab-mapping-excel1"
          onClick={() => setActiveTab('excel1')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between gap-2 shadow-xs ${
            activeTab === 'excel1'
              ? 'border-blue-500 bg-blue-50/40 ring-2 ring-blue-500/20'
              : 'border-slate-200 bg-white hover:bg-slate-50'
          }`}
        >
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1 text-[10px] font-bold uppercase tracking-wider text-blue-700 bg-blue-100 px-2 py-0.5 rounded">
              <FileInput className="w-3 h-3" />
              Section A
            </span>
            <span className={`text-[10px] px-2 py-0.5 rounded font-bold ${getTabStatus('excel1').bg}`}>
              {getTabStatus('excel1').text}
            </span>
          </div>
          <div>
            <h4 className="text-xs font-bold text-slate-900">EXCEL 1 — SOURCE</h4>
            <p className="text-[11px] text-slate-500">5 Source Input Fields</p>
          </div>
        </button>

        {/* Tab B: Excel 2 */}
        <button
          id="tab-mapping-excel2"
          onClick={() => setActiveTab('excel2')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between gap-2 shadow-xs ${
            activeTab === 'excel2'
              ? 'border-indigo-500 bg-indigo-50/40 ring-2 ring-indigo-500/20'
              : 'border-slate-200 bg-white hover:bg-slate-50'
          }`}
        >
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1 text-[10px] font-bold uppercase tracking-wider text-indigo-700 bg-indigo-100 px-2 py-0.5 rounded">
              <FileInput className="w-3 h-3" />
              Section B
            </span>
            <span className={`text-[10px] px-2 py-0.5 rounded font-bold ${getTabStatus('excel2').bg}`}>
              {getTabStatus('excel2').text}
            </span>
          </div>
          <div>
            <h4 className="text-xs font-bold text-slate-900">EXCEL 2 — SOURCE</h4>
            <p className="text-[11px] text-slate-500">6 Source Input Fields</p>
          </div>
        </button>

        {/* Tab C: Excel 3 */}
        <button
          id="tab-mapping-excel3"
          onClick={() => setActiveTab('excel3')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between gap-2 shadow-xs ${
            activeTab === 'excel3'
              ? 'border-emerald-500 bg-emerald-50/40 ring-2 ring-emerald-500/20'
              : 'border-slate-200 bg-white hover:bg-slate-50'
          }`}
        >
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1 text-[10px] font-bold uppercase tracking-wider text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded">
              <FileOutput className="w-3 h-3" />
              Section C
            </span>
            <span className={`text-[10px] px-2 py-0.5 rounded font-bold ${getTabStatus('excel3').bg}`}>
              {getTabStatus('excel3').text}
            </span>
          </div>
          <div>
            <h4 className="text-xs font-bold text-slate-900">EXCEL 3 — TARGET</h4>
            <p className="text-[11px] text-slate-500">9 Output Targets + Meeting</p>
          </div>
        </button>

        {/* View All Sections */}
        <button
          id="tab-mapping-all"
          onClick={() => setActiveTab('all')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between gap-2 shadow-xs ${
            activeTab === 'all'
              ? 'border-purple-500 bg-purple-50/40 ring-2 ring-purple-500/20'
              : 'border-slate-200 bg-white hover:bg-slate-50'
          }`}
        >
          <div className="flex items-center justify-between">
            <span className="inline-flex items-center gap-1 text-[10px] font-bold uppercase tracking-wider text-purple-700 bg-purple-100 px-2 py-0.5 rounded">
              <Layers className="w-3 h-3" />
              All Sections
            </span>
            <span className="text-[10px] text-slate-500 font-medium">Overview</span>
          </div>
          <div>
            <h4 className="text-xs font-bold text-slate-900">ALL SECTIONS (A + B + C)</h4>
            <p className="text-[11px] text-slate-500">View complete mapping sequence</p>
          </div>
        </button>
      </div>

      {/* Main Content Area based on selected Tab */}
      <div className="space-y-8">
        {(activeTab === 'excel1' || activeTab === 'all') && renderSectionA()}
        {(activeTab === 'excel2' || activeTab === 'all') && renderSectionB()}
        {(activeTab === 'excel3' || activeTab === 'all') && renderSectionC()}
      </div>

      {/* Validation Feedback Banner */}
      {validationResult && (
        <div
          id="mapping-validation-banner"
          className={`p-4 rounded-xl border text-xs shadow-xs ${
            validationResult.isValid
              ? 'bg-emerald-50 border-emerald-300 text-emerald-900'
              : 'bg-rose-50 border-rose-300 text-rose-900'
          }`}
        >
          <div className="flex items-center gap-2 font-bold mb-1 text-sm">
            {validationResult.isValid ? (
              <>
                <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                <span>Mapping Validated Successfully!</span>
              </>
            ) : (
              <>
                <AlertCircle className="w-5 h-5 text-rose-600" />
                <span>Mapping Validation Incomplete</span>
              </>
            )}
          </div>
          {validationResult.isValid ? (
            <p className="text-emerald-700">
              All 5 source fields in Excel 1, 6 source fields in Excel 2, and required target columns in Excel 3 are valid. You may proceed to Step 3.
            </p>
          ) : (
            <ul className="list-disc pl-5 space-y-1 mt-2 text-rose-800">
              {validationResult.errors.map((e, idx) => (
                <li key={idx}>{e}</li>
              ))}
            </ul>
          )}
        </div>
      )}

      {/* Sticky Action Footer */}
      <div className="sticky bottom-0 z-40 bg-white/95 backdrop-blur-md rounded-xl p-4 border border-slate-200 shadow-lg flex flex-wrap items-center justify-between gap-4">
        {/* Mapping Status summary */}
        <div className="flex items-center gap-4">
          <span className="text-xs font-bold text-slate-700">Mapping Architecture:</span>
          <div className="flex flex-wrap items-center gap-2.5">
            {[
              { key: 'excel1' as WorkbookType, label: 'Sec A: Excel 1' },
              { key: 'excel2' as WorkbookType, label: 'Sec B: Excel 2' },
              { key: 'excel3' as WorkbookType, label: 'Sec C: Excel 3' },
            ].map((item) => {
              const valid = mappings[item.key]?.isValid && (item.key !== 'excel3' || excel3Dupes.size === 0);
              return (
                <span
                  key={item.key}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium border ${
                    valid
                      ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                      : 'bg-amber-50 text-amber-700 border-amber-200'
                  }`}
                >
                  {valid ? <CheckCircle2 className="w-3.5 h-3.5" /> : <AlertTriangle className="w-3.5 h-3.5" />}
                  <span>
                    {item.label}: {valid ? 'Valid' : 'Needs Mapping'}
                  </span>
                </span>
              );
            })}
          </div>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-3">
          <button
            onClick={onBackToStep1}
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition"
          >
            ← Back to Step 1
          </button>

          <button
            id="btn-validate-mapping"
            onClick={handleValidateMapping}
            disabled={isValidating}
            className="px-4 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition flex items-center gap-2"
          >
            {isValidating && <RefreshCw className="w-3.5 h-3.5 animate-spin" />}
            Validate Mapping
          </button>

          <button
            id="btn-proceed-step3"
            onClick={onProceedToStep3}
            disabled={!canProceed}
            className={`px-5 py-2 text-xs font-bold rounded-lg transition flex items-center gap-2 shadow-xs ${
              canProceed
                ? 'bg-blue-600 hover:bg-blue-700 text-white cursor-pointer'
                : 'bg-slate-200 text-slate-400 border border-slate-300 cursor-not-allowed'
            }`}
            title={
              !canProceed
                ? 'All required source fields (Excel 1, Excel 2) and output targets (Excel 3) must be mapped'
                : 'Proceed to Step 3: Select Jobs'
            }
          >
            <span>Proceed to Step 3 →</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  );
};
