import jSuites from 'jsuites';
import { getColumnNameFromId } from './internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';
import type { MaskOptions } from '../types/jsuites';

// The value that a cell shows: the result of its formula (worksheet.formulas calculates it), or its value with the mask of
// its column. These functions read the worksheet and write no cell, so the modules that change cells can import them
// without an import cycle.

/**
 * Get the column whose options a cell uses. The first cell of a merged range whose first column has the
 * type hidden uses the options of the first visible column of the range, because it shows the range.
 * A checkbox or radio column is not used, because these types write a boolean to the data of the cell.
 */
export const getCellColumnIndex = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);

    if (obj.options.columns && obj.options.columns[x] && obj.options.columns[x].type == 'hidden' && obj.options.mergeCells) {
        const merged = obj.options.mergeCells[getColumnNameFromId([x, y])];
        if (merged) {
            for (let i = x + 1; i < x + merged[0]; i++) {
                if (obj.headers[i] && obj.headers[i].style.display != 'none' && (!obj.options.columns[i] || obj.options.columns[i].type != 'hidden')) {
                    if (obj.options.columns[i] && (obj.options.columns[i].type == 'checkbox' || obj.options.columns[i].type == 'radio')) {
                        break;
                    }
                    return i;
                }
            }
        }
    }

    return x;
};

/**
 * The value that a cell shows: the result of a formula, or the value with the mask of the column. The DOM converts a
 * number, a boolean or undefined to text.
 */
export const parseValue = function (this: Worksheet, i: number, j: number, value: jspreadsheet.CellValue | undefined, cell?: HTMLElement) {
    const obj = this;

    if (('' + value).substr(0, 1) == '=' && obj.parent.config.parseFormulas != false) {
        // The cell shows the result as text, as the DOM converts it. An array from a range shows its items with commas.
        value = obj.formulas.execute(value as string, i, j) as jspreadsheet.CellValue;
    }

    // Column options
    const options = obj.options.columns && obj.options.columns[getCellColumnIndex.call(obj, i, j)];
    if (options && !isFormula(value)) {
        // Mask options
        let opt = null;
        if ((opt = getMask(options))) {
            if (value && value == Number(value)) {
                value = Number(value);
            }
            // Process the decimals to match the mask
            let masked = jSuites.mask.render(value, opt, true);
            // Negative indication
            if (cell) {
                if (opt.mask) {
                    const t = opt.mask.split(';');
                    if (t[1]) {
                        const t1 = t[1].match(new RegExp('\\[Red\\]', 'gi'));
                        if (t1) {
                            // A text compares as a number
                            if ((value as number) < 0) {
                                cell.classList.add('red');
                            } else {
                                cell.classList.remove('red');
                            }
                        }
                        const t2 = t[1].match(new RegExp('\\(', 'gi'));
                        if (t2) {
                            if ((value as number) < 0) {
                                masked = '(' + masked + ')';
                            }
                        }
                    }
                }
            }

            if (masked) {
                value = masked;
            }
        }
    }

    return value;
};

/**
 * The value is a formula
 */
export const isFormula = function (value: unknown) {
    const v = ('' + value)[0];
    return v == '=' || v == '#' ? true : false;
};

/**
 * Get the mask in the jSuites.mask format
 */
export const getMask = function (o: Worksheet['options']['columns'][number]): MaskOptions | null {
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
};
