import type { Worksheet } from '../types/internal';

/**
 * Apply a change of rows or columns that another source made, for example a server that the worksheet shows. During the
 * change, the worksheet writes no history record and dispatches no event, the insert and the delete ask no question and
 * ignore the allow options, and the formulas keep their texts.
 * @param worksheet - The worksheet.
 * @param change - The change.
 * @returns The result of the change.
 */
export function applyQuietly<T>(worksheet: Worksheet, change: () => T): T {
    const spreadsheet = worksheet.parent;
    const ignoreEvents = spreadsheet.ignoreEvents;
    const ignoreHistory = worksheet.history.ignore;
    const quiet = worksheet.quiet;

    spreadsheet.ignoreEvents = true;
    worksheet.history.ignore = true;
    worksheet.quiet = true;
    try {
        return change();
    } finally {
        spreadsheet.ignoreEvents = ignoreEvents;
        worksheet.history.ignore = ignoreHistory;
        worksheet.quiet = quiet;
    }
}
