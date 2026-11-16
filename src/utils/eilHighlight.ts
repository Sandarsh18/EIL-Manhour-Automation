/**
 * Highlighting and Value Verification Helpers for EIL Manhour Automation & Audit
 * 
 * Strict rules:
 * - isPresentValue distinguishes 0, "0", 0.0, "0.0", null, undefined, "", "   "
 *   from actual populated values (1, 4, 3, 14, 12.5, 65.5, 0.5, etc.)
 * - Zero and blank are NOT considered populated.
 * - Non-zero numeric values ARE considered populated and receive amber audit highlighting.
 * - Decimals (0.5, 7.5, 12.5, 65.5) are preserved exactly.
 */

/**
 * Determines whether a value represents actual, populated numeric data (> 0 or non-zero).
 * Correctly distinguishes 0, "0", 0.0, "0.0", null, undefined, "", "   " from actual values.
 */
export function isPresentValue(val: unknown): boolean {
  if (val === null || val === undefined) return false;

  if (typeof val === 'number') {
    if (isNaN(val)) return false;
    return val !== 0;
  }

  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '') return false;
    const num = Number(trimmed);
    if (isNaN(num)) return false;
    return num !== 0;
  }

  return false;
}

/**
 * Formats a value for display in the table or audit cards.
 * Preserves decimal places accurately (e.g. 12.5, 65.5, 0.5).
 */
export function formatDisplayValue(val: unknown, fallback: string = '0'): string {
  if (val === null || val === undefined) return fallback;
  
  if (typeof val === 'number') {
    if (isNaN(val)) return fallback;
    // If integer, return as integer string; if decimal, preserve decimals up to 2 places without trailing zeros
    if (Number.isInteger(val)) {
      return String(val);
    }
    return String(Math.round(val * 100) / 100);
  }

  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '') return fallback;
    const num = Number(trimmed);
    if (!isNaN(num)) {
      if (Number.isInteger(num)) return String(num);
      return String(Math.round(num * 100) / 100);
    }
    return trimmed;
  }

  return String(val);
}

/**
 * Returns Tailwind CSS class string for a table cell container based on whether
 * the value is an actual populated non-zero number vs. a neutral zero or blank.
 */
export function getValueBadgeClass(val: unknown, isTotal: boolean = false): string {
  const isPopulated = isPresentValue(val);

  if (isTotal) {
    if (isPopulated) {
      return 'inline-flex items-center justify-end min-w-[46px] px-2.5 py-0.5 rounded-md bg-amber-100/90 border border-amber-400 text-blue-950 font-bold font-mono text-xs shadow-2xs';
    }
    return 'inline-flex items-center justify-end min-w-[46px] px-2 py-0.5 text-slate-400 font-mono text-xs';
  }

  if (isPopulated) {
    return 'inline-flex items-center justify-end min-w-[34px] px-2 py-0.5 rounded-md bg-amber-50/90 border border-amber-300 text-amber-950 font-semibold font-mono text-xs shadow-2xs transition-colors';
  }

  if (val === 0 || val === '0' || val === 0.0 || val === '0.0') {
    return 'inline-flex items-center justify-end min-w-[28px] px-2 py-0.5 text-slate-400 font-mono text-xs';
  }

  return 'inline-flex items-center justify-end min-w-[28px] px-2 py-0.5 text-slate-300 font-mono text-xs';
}
