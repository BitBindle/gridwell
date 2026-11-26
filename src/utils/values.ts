import type { Worksheet } from '../types/internal';
import type { MaskOptions } from '../types/jsuites';

// Whether a value is a formula, and the mask of a column. worksheet.cells uses them for the value that a cell shows, and
// the editor for the value that it edits.

/**
 * The value is a formula
 */
export function isFormula(value: unknown) {
    const v = `${value}`[0];
    return v == '=' || v == '#' ? true : false;
}

/**
 * Get the mask in the jSuites.mask format
 */
export function getMask(o: Worksheet['options']['columns'][number]): MaskOptions | null {
    if (o.format || o.mask || o.locale) {
        const opt: MaskOptions = {};
        if (o.mask) {
            opt.mask = o.mask;
        } else if (o.format) {
            opt.mask = o.format;
        } else {
            opt.locale = o.locale;
            opt.options = o.options;
        }

        if (o.decimal) {
            if (!opt.options) {
                opt.options = {};
            }
            opt.options = { decimal: o.decimal };
        }
        return opt;
    }

    return null;
}
