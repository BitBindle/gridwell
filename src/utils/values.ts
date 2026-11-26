import type { Worksheet } from '../types/internal';
import type { Mask } from './mask.js';

// Whether a value is a formula, and the mask of a column. worksheet.cells uses them for the value that a cell shows, and
// the editor for the value that it edits.

/**
 * The value is a formula
 */
export function isFormula(value: unknown) {
    const v = `${value}`[0];
    return v === '=' || v === '#' ? true : false;
}

/**
 * The mask of a column: its mask or format, or its locale with the options of Intl.NumberFormat, and its decimal
 * separator.
 * @param o - The column.
 * @returns The mask, or null when the column has none.
 */
export function getMask(o: Worksheet['options']['columns'][number]): Mask | null {
    if (o.format || o.mask || o.locale) {
        const opt: Mask = {};
        if (o.mask) {
            opt.mask = o.mask;
        } else if (o.format) {
            opt.mask = o.format;
        } else {
            opt.locale = o.locale;
            // A column with a locale is a numeric column, whose options are those of Intl.NumberFormat
            opt.options = o.options as Intl.NumberFormatOptions | undefined;
        }

        if (o.decimal) {
            opt.decimal = o.decimal;
        }
        return opt;
    }

    return null;
}
