/**
 * Centralized Date Parser for EIL Manhour Automation & Audit
 * 
 * Supports:
 * - Excel serial dates (handles the historic Excel 1900 leap year bug)
 * - Native Date objects
 * - ISO date strings (e.g. 2026-08-15)
 * - Indian/UK format strings (e.g. DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY)
 * - US format strings (e.g. MM/DD/YYYY)
 * - OCS Six-Calendar-Month window calculation
 */

export interface ParsedDateResult {
  isValid: boolean;
  date: Date | null;
  isoString: string | null;
  formattedDate: string | null; // YYYY-MM-DD
  rawInput: any;
  error?: string;
}

const MONTH_NAME_MAP: Record<string, number> = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

export class EilDateParser {
  /**
   * Parses various Excel date representations into a unified JavaScript Date.
   */
  public static parse(val: any): ParsedDateResult {
    if (val === undefined || val === null || val === '') {
      return {
        isValid: false,
        date: null,
        isoString: null,
        formattedDate: null,
        rawInput: val,
        error: 'Date value is empty or null',
      };
    }

    // Already a Date object
    if (val instanceof Date) {
      if (isNaN(val.getTime())) {
        return {
          isValid: false,
          date: null,
          isoString: null,
          formattedDate: null,
          rawInput: val,
          error: 'Invalid Date object',
        };
      }
      return this.successResult(val, val);
    }

    // Excel serial number (numeric value, usually between 1000 and 100000)
    if (typeof val === 'number' && !isNaN(val)) {
      const parsedDate = this.fromExcelSerial(val);
      if (parsedDate && !isNaN(parsedDate.getTime())) {
        return this.successResult(parsedDate, val);
      }
      return {
        isValid: false,
        date: null,
        isoString: null,
        formattedDate: null,
        rawInput: val,
        error: `Could not convert Excel serial number ${val} to date`,
      };
    }

    // String date
    if (typeof val === 'string') {
      const trimmed = val.trim();
      if (!trimmed) {
        return {
          isValid: false,
          date: null,
          isoString: null,
          formattedDate: null,
          rawInput: val,
          error: 'Empty string date',
        };
      }

      // Check if string is a numeric string (Excel serial exported as string)
      if (/^\d{4,6}(\.\d+)?$/.test(trimmed)) {
        const num = parseFloat(trimmed);
        const parsedDate = this.fromExcelSerial(num);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          return this.successResult(parsedDate, val);
        }
      }

      // Format: D-MMM-YY, DD-MMM-YY, D-MMM-YYYY, DD-MMM-YYYY (e.g. 7-Oct-26, 1-Apr-26, 9-Oct-25, 28-Oct-2026)
      const dMmmMatch = trimmed.match(/^(\d{1,2})[\s\-\/\.]*([a-zA-Z]{3,9})[\s\-\/\.]*(\d{2,4})$/);
      if (dMmmMatch) {
        const day = parseInt(dMmmMatch[1], 10);
        const monKey = dMmmMatch[2].toLowerCase().slice(0, 3);
        const month = MONTH_NAME_MAP[monKey];
        let year = parseInt(dMmmMatch[3], 10);
        if (dMmmMatch[3].length === 2) {
          year = year < 50 ? 2000 + year : 1900 + year;
        }
        if (month && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const d = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
          if (d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day) {
            return this.successResult(d, val);
          }
        }
      }

      // DD/MM/YY or DD-MM-YY (Indian 2-digit year format: e.g. 12/10/26 -> 2026-10-12)
      const dmy2Match = trimmed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2})$/);
      if (dmy2Match) {
        const day = parseInt(dmy2Match[1], 10);
        const month = parseInt(dmy2Match[2], 10);
        const yy = parseInt(dmy2Match[3], 10);
        const year = yy < 50 ? 2000 + yy : 1900 + yy;
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const d = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
          if (d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day) {
            return this.successResult(d, val);
          }
        }
      }

      // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
      const dmyMatch = trimmed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
      if (dmyMatch) {
        const day = parseInt(dmyMatch[1], 10);
        const month = parseInt(dmyMatch[2], 10);
        const year = parseInt(dmyMatch[3], 10);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const d = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
          if (d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day) {
            return this.successResult(d, val);
          }
        }
      }

      // YYYY-MM-DD or YYYY/MM/DD
      const ymdMatch = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (ymdMatch) {
        const year = parseInt(ymdMatch[1], 10);
        const month = parseInt(ymdMatch[2], 10);
        const day = parseInt(ymdMatch[3], 10);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const d = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
          if (d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day) {
            return this.successResult(d, val);
          }
        }
      }

      // Fallback to Date.parse, constructing clean UTC date to avoid timezone shift
      const fallbackTime = Date.parse(trimmed);
      if (!isNaN(fallbackTime)) {
        const localD = new Date(fallbackTime);
        const d = new Date(Date.UTC(localD.getFullYear(), localD.getMonth(), localD.getDate(), 0, 0, 0));
        return this.successResult(d, val);
      }

      return {
        isValid: false,
        date: null,
        isoString: null,
        formattedDate: null,
        rawInput: val,
        error: `Unrecognized date string format: "${trimmed}"`,
      };
    }

    return {
      isValid: false,
      date: null,
      isoString: null,
      formattedDate: null,
      rawInput: val,
      error: `Unsupported value type: ${typeof val}`,
    };
  }

  /**
   * Convert an Excel serial number to a JavaScript Date.
   * Note: Excel's 1900 date system has a known leap year bug where 1900-02-29 is considered day 60.
   */
  public static fromExcelSerial(serial: number): Date | null {
    if (serial < 1) return null;
    // Excel Day 1 = Jan 1, 1900
    // Adjust for Lotus 1-2-3 1900 leap-year bug (day 60 was Feb 29 1900, which didn't exist)
    const days = serial > 60 ? serial - 2 : serial - 1;
    const epoch = new Date(Date.UTC(1900, 0, 1, 0, 0, 0));
    const msPerDay = 24 * 60 * 60 * 1000;
    return new Date(epoch.getTime() + Math.round(days * msPerDay));
  }

  /**
   * Computes the exact evaluation month OCS eligibility window:
   * 1st of evaluation month to the last day of the evaluation month.
   * 
   * Example:
   * Evaluation month = "2026-08"
   * Window:
   * Start: 2026-08-01 00:00:00 UTC
   * End: 2026-08-31 23:59:59.999 UTC
   */
  public static getOcsEligibilityWindow(evaluationMonthYyyyMm: string): {
    start: Date;
    end: Date;
    startFormatted: string;
    endFormatted: string;
    monthsIncluded: string[];
  } {
    const window = this.getEvaluationMonthWindow(evaluationMonthYyyyMm);
    return {
      start: window.start,
      end: window.end,
      startFormatted: window.startFormatted,
      endFormatted: window.endFormatted,
      monthsIncluded: [evaluationMonthYyyyMm],
    };
  }

  /**
   * Evaluates if an OCS Date falls within the active evaluation month.
   * Eligible OCS Date range for 2026-08: 2026-08-01 through 2026-08-31.
   */
  public static isOcsDateInEvaluationMonth(dateVal: any, evaluationMonthYyyyMm: string): boolean {
    return this.isDateInEvaluationMonth(dateVal, evaluationMonthYyyyMm);
  }

  /**
   * Determines if a date is within the evaluation month (current month + current year).
   */
  public static isDateInEvaluationMonth(dateVal: any, evaluationMonthYyyyMm: string): boolean {
    const parsed = this.parse(dateVal);
    if (!parsed.isValid || !parsed.date) return false;

    const window = this.getEvaluationMonthWindow(evaluationMonthYyyyMm);
    const t = parsed.date.getTime();
    return t >= window.start.getTime() && t <= window.end.getTime();
  }

  /**
   * Computes the active evaluation month boundary: 1st of month to end of month.
   */
  public static getEvaluationMonthWindow(evaluationMonthYyyyMm: string): {
    start: Date;
    end: Date;
    year: number;
    month: number;
    daysInMonth: number;
    startFormatted: string;
    endFormatted: string;
  } {
    const match = evaluationMonthYyyyMm.trim().match(/^(\d{4})-(\d{2})$/);
    if (!match) {
      throw new Error(`Invalid evaluation month format: ${evaluationMonthYyyyMm}. Expected YYYY-MM.`);
    }

    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10); // 1 to 12

    const lastDayOfMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
    const start = new Date(Date.UTC(year, month - 1, 1, 0, 0, 0, 0));
    const end = new Date(Date.UTC(year, month - 1, lastDayOfMonth, 23, 59, 59, 999));

    const pad = (n: number) => String(n).padStart(2, '0');
    return {
      start,
      end,
      year,
      month,
      daysInMonth: lastDayOfMonth,
      startFormatted: `${year}-${pad(month)}-01`,
      endFormatted: `${year}-${pad(month)}-${pad(lastDayOfMonth)}`,
    };
  }

  /**
   * Evaluates Inspection attendance dates against the active evaluation month.
   * Handles:
   * - Records completely inside current month/year -> INCLUDED (full authoritative active days)
   * - Records spanning across month boundary -> INCLUDED (full authoritative active days; never pro-rated)
   * - Records completely outside current month/year -> EXCLUDED (0 days, ignored historical)
   */
  public static evaluateInspectionDates(
    rawFrom: any,
    rawUpto: any,
    evaluationMonthYyyyMm: string,
    totalReportedDays: number = 0
  ): {
    isIncluded: boolean;
    activeDays: number;
    status: 'FULL' | 'EXCLUDED' | 'INVALID';
    reason: string;
    fromParsed: ParsedDateResult;
    uptoParsed: ParsedDateResult;
  } {
    let fromParsed = this.parse(rawFrom);
    let uptoParsed = this.parse(rawUpto);

    // If only one attendance date is valid, anchor both bounds to that valid date
    if (fromParsed.isValid && fromParsed.date && (!uptoParsed.isValid || !uptoParsed.date)) {
      uptoParsed = fromParsed;
    } else if (uptoParsed.isValid && uptoParsed.date && (!fromParsed.isValid || !fromParsed.date)) {
      fromParsed = uptoParsed;
    }

    if (!fromParsed.isValid || !fromParsed.date || !uptoParsed.isValid || !uptoParsed.date) {
      return {
        isIncluded: false,
        activeDays: 0,
        status: 'INVALID',
        reason: 'One or both inspection attendance dates are invalid or unparseable.',
        fromParsed,
        uptoParsed,
      };
    }

    const window = this.getEvaluationMonthWindow(evaluationMonthYyyyMm);
    const fTime = Math.min(fromParsed.date.getTime(), uptoParsed.date.getTime());
    const uTime = Math.max(fromParsed.date.getTime(), uptoParsed.date.getTime());
    const wStart = window.start.getTime();
    const wEnd = window.end.getTime();

    // Completely outside the evaluation month
    if (uTime < wStart || fTime > wEnd) {
      return {
        isIncluded: false,
        activeDays: 0,
        status: 'EXCLUDED',
        reason: `Inspection dates (${fromParsed.formattedDate} to ${uptoParsed.formattedDate}) fall outside evaluation period (${evaluationMonthYyyyMm}). Historical call records ignored.`,
        fromParsed,
        uptoParsed,
      };
    }

    // Completely inside evaluation month
    if (fTime >= wStart && uTime <= wEnd) {
      return {
        isIncluded: totalReportedDays > 0,
        activeDays: totalReportedDays,
        status: 'FULL',
        reason: `Inspection dates fully within ${evaluationMonthYyyyMm}.`,
        fromParsed,
        uptoParsed,
      };
    }

    // Overlaps evaluation month boundary (e.g. 28-Sep -> 01-Oct): preserve full authoritative No. of Days
    return {
      isIncluded: totalReportedDays > 0,
      activeDays: totalReportedDays,
      status: 'FULL',
      reason: `Inspection period overlaps ${evaluationMonthYyyyMm} (${fromParsed.formattedDate} to ${uptoParsed.formattedDate}). Full authoritative ${totalReportedDays} days applied.`,
      fromParsed,
      uptoParsed,
    };
  }

  /**
   * Evaluates QAP approval date against the active evaluation month and year.
   * Only records whose QAP approval date falls strictly within the active evaluation month
   * and year are included. Historical records from older years or other months are excluded.
   */
  public static evaluateQapDate(
    rawDate: any,
    evaluationMonthYyyyMm: string
  ): { isIncluded: boolean; reason?: string; parsedDate?: string } {
    if (rawDate === undefined || rawDate === null || String(rawDate).trim() === '') {
      return {
        isIncluded: false,
        reason: `QAP record has no date in evaluation cycle (expected ${evaluationMonthYyyyMm}). Excluded.`,
      };
    }

    const parsed = this.parse(rawDate);
    if (!parsed.isValid || !parsed.date) {
      return {
        isIncluded: false,
        reason: `QAP date "${rawDate}" is invalid or unparseable. Excluded from ${evaluationMonthYyyyMm}.`,
      };
    }

    const inMonth = this.isDateInEvaluationMonth(parsed.date, evaluationMonthYyyyMm);
    if (!inMonth) {
      return {
        isIncluded: false,
        reason: `QAP date ${parsed.formattedDate} is outside active evaluation cycle (${evaluationMonthYyyyMm}). Historical QAP record excluded.`,
        parsedDate: parsed.formattedDate || undefined,
      };
    }

    return {
      isIncluded: true,
      reason: `QAP date ${parsed.formattedDate} is within active evaluation cycle (${evaluationMonthYyyyMm}).`,
      parsedDate: parsed.formattedDate || undefined,
    };
  }

  /**
   * Evaluates Review dates against the active evaluation month and year.
   * Supports single dates, Excel serials, native Date objects, and multi-date
   * composite strings (e.g. "1,9,21/10/26", "22-Oct-26", "13-10-2026; 23/10/2026").
   */
  public static evaluateReviewDate(
    rawVal: any,
    evaluationMonthYyyyMm: string
  ): { isIncluded: boolean; reason?: string; parsedDate?: string } {
    if (rawVal === undefined || rawVal === null || String(rawVal).trim() === '') {
      return {
        isIncluded: false,
        reason: `Review dates record has no date in evaluation cycle (expected ${evaluationMonthYyyyMm}). Excluded.`,
      };
    }

    const isComposite =
      typeof rawVal === 'string' && (rawVal.includes(',') || rawVal.includes(';') || rawVal.includes('&'));

    // Try direct single-date parsing only for non-composite date representations
    if (!isComposite) {
      const parsed = this.parse(rawVal);
      if (parsed.isValid && parsed.date) {
        const inMonth = this.isDateInEvaluationMonth(parsed.date, evaluationMonthYyyyMm);
        if (inMonth) {
          return {
            isIncluded: true,
            reason: `Review date ${parsed.formattedDate} is within active evaluation cycle (${evaluationMonthYyyyMm}).`,
            parsedDate: parsed.formattedDate || undefined,
          };
        } else {
          return {
            isIncluded: false,
            reason: `Review date ${parsed.formattedDate} is outside active evaluation cycle (${evaluationMonthYyyyMm}). Historical Review record excluded.`,
            parsedDate: parsed.formattedDate || undefined,
          };
        }
      }
    }

    // Handle composite strings with multiple dates or date fragments
    const str = String(rawVal).trim();
    const dateRegexes = [
      /\b\d{4}[\-\/]\d{1,2}[\-\/]\d{1,2}\b/g,
      /\b\d{1,2}[\s\-\/\.]*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-\/\.]*\d{2,4}\b/gi,
      /\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}\b/g,
      /\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2}\b/g,
    ];

    const foundCandidates: string[] = [];
    for (const rx of dateRegexes) {
      const matches = str.match(rx);
      if (matches) {
        foundCandidates.push(...matches);
      }
    }

    for (const cand of foundCandidates) {
      const candParsed = this.parse(cand);
      if (candParsed.isValid && candParsed.date) {
        if (this.isDateInEvaluationMonth(candParsed.date, evaluationMonthYyyyMm)) {
          return {
            isIncluded: true,
            reason: `Review date ${candParsed.formattedDate} in composite string is within active evaluation cycle (${evaluationMonthYyyyMm}).`,
            parsedDate: candParsed.formattedDate || undefined,
          };
        }
      }
    }

    if (foundCandidates.length > 0) {
      return {
        isIncluded: false,
        reason: `Review dates (${foundCandidates.join(', ')}) are outside active evaluation cycle (${evaluationMonthYyyyMm}).`,
      };
    }

    // Shorthand checks like "1,9,21/10/26" or "16,17-Oct-26":
    const segments = str.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
    let suffixMonYear = '';
    for (let i = segments.length - 1; i >= 0; i--) {
      const monMatch =
        segments[i].match(/(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-\/\.]*\d{2,4}/i) ||
        segments[i].match(/[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/);
      if (monMatch) {
        suffixMonYear = monMatch[0];
        break;
      }
    }

    if (suffixMonYear) {
      for (const seg of segments) {
        const testDateStr = seg.includes(suffixMonYear)
          ? seg
          : `${seg}${suffixMonYear.startsWith('-') || suffixMonYear.startsWith('/') ? '' : '-'}${suffixMonYear}`;
        const p = this.parse(testDateStr);
        if (p.isValid && p.date && this.isDateInEvaluationMonth(p.date, evaluationMonthYyyyMm)) {
          return {
            isIncluded: true,
            reason: `Review date ${p.formattedDate} is within active evaluation cycle (${evaluationMonthYyyyMm}).`,
            parsedDate: p.formattedDate || undefined,
          };
        }
      }
    }

    return {
      isIncluded: false,
      reason: `Review dates "${rawVal}" is unparseable or outside active evaluation cycle (${evaluationMonthYyyyMm}).`,
    };
  }

  /**
   * Determines if a date is within the OCS eligibility window (inclusive).
   */
  public static isDateInOcsWindow(dateVal: any, evaluationMonthYyyyMm: string): boolean {
    const parsed = this.parse(dateVal);
    if (!parsed.isValid || !parsed.date) return false;

    const window = this.getOcsEligibilityWindow(evaluationMonthYyyyMm);
    const t = parsed.date.getTime();
    return t >= window.start.getTime() && t <= window.end.getTime();
  }

  private static successResult(d: Date, raw: any): ParsedDateResult {
    const pad = (n: number) => String(n).padStart(2, '0');
    const formatted = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    return {
      isValid: true,
      date: d,
      isoString: d.toISOString(),
      formattedDate: formatted,
      rawInput: raw,
    };
  }
}
