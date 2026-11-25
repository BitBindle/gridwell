import jSuites from 'jsuites';
import formula from '@jspreadsheet/formula';
import dispatch from './dispatch.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, WorksheetMethod } from '../types/internal';
import type { MaskOptions } from '../types/jsuites';

// The value that a cell shows: the result of its formula, or its value with the mask of its column. These functions read
// the worksheet and write no cell, so the modules that change cells can import them without an import cycle.

/**
 * Trying to extract a number from a string
 */
const parseNumber = function (this: Worksheet, value: unknown, columnNumber: number) {
    const obj = this;

    // Decimal point. Column 0 has one too.
    const column = obj.options.columns[columnNumber];
    const decimal = column && column.decimal ? column.decimal : '.';

    // Parse both parts of the number
    // Text, then its parts, then the digits of the integer part, then the digits as text
    let number: string | (string | string[] | null)[] = '' + value;
    number = number.split(decimal);
    number[0] = (number[0] as string).match(/[+-]?[0-9]/g);
    if (number[0]) {
        number[0] = (number[0] as string[]).join('');
    }
    if (number[1]) {
        // The pattern also matches an empty text, so there is a match
        number[1] = (number[1] as string).match(/[0-9]*/g)!.join('');
    }

    // Is a valid number
    if (number[0] && Number.isInteger(Number(number[0]))) {
        if (!number[1]) {
            value = Number(number[0] + '.00');
        } else {
            value = Number(number[0] + '.' + number[1]);
        }
    } else {
        value = null;
    }

    return value;
};

/**
 * Parse formulas
 */
export const executeFormula: WorksheetMethod<'executeFormula'> = function (expression, x, y) {
    const obj = this;

    // The results and the formulas in progress, by cell name
    const formulaResults: Record<string, unknown> = {};
    const formulaLoopProtection: Record<string, boolean> = {};

    // A formula engine of the application calculates the formula in place of the formula plugin
    const formulaEngine = typeof obj.parent.config.formulaEngine === 'function' ? obj.parent.config.formulaEngine : null;

    // Execute formula with loop protection
    const execute = function (expression: string, x?: number, y?: number) {
        // The formula text as typed, for the formula engine
        const text = expression;

        // Parent column identification. executeFormula without x and y runs a formula that is in no cell.
        const parentId = getColumnNameFromId([x!, y!]);

        // Code protection
        if (formulaLoopProtection[parentId]) {
            console.error('Reference loop detected');
            return '#ERROR';
        }

        formulaLoopProtection[parentId] = true;

        // Convert range tokens
        const tokensUpdate = function (tokens: string[]) {
            for (let index = 0; index < tokens.length; index++) {
                const f = [];
                const token = tokens[index].split(':');
                const e1 = getIdFromColumnName(token[0], true);
                const e2 = getIdFromColumnName(token[1], true);

                let x1, x2;

                if (e1[0] <= e2[0]) {
                    x1 = e1[0];
                    x2 = e2[0];
                } else {
                    x1 = e2[0];
                    x2 = e1[0];
                }

                let y1, y2;

                if (e1[1] <= e2[1]) {
                    y1 = e1[1];
                    y2 = e2[1];
                } else {
                    y1 = e2[1];
                    y2 = e1[1];
                }

                for (let j = y1; j <= y2; j++) {
                    for (let i = x1; i <= x2; i++) {
                        f.push(getColumnNameFromId([i, j]));
                    }
                }

                expression = expression.replace(tokens[index], f.join(','));
            }
        };

        // Range with $ remove $ (string literals are kept untouched)
        expression = expression.replace(/("(?:[^"]|"")*")|\$?([A-Z]+)\$?([0-9]+)/g, function (match, literal, column, row) {
            return literal ? literal : column + row;
        });

        let tokens = expression.match(/([A-Z]+[0-9]+):([A-Z]+[0-9]+)/g);
        if (tokens && tokens.length) {
            tokensUpdate(tokens);
        }

        // Get tokens
        tokens = expression.match(/([A-Z]+[0-9]+)/g);

        // Direct self-reference protection
        if (tokens && tokens.indexOf(parentId) > -1) {
            console.error('Self Reference detected');
            return '#ERROR';
        } else {
            // Expressions to be used in the parsing
            const formulaExpressions: Record<string, unknown> = {};

            if (tokens) {
                for (let i = 0; i < tokens.length; i++) {
                    // Keep chain
                    obj.formulas.add(tokens[i], parentId);

                    // Do not calculate again. The formula engine reads the values of the references itself.
                    if (!formulaEngine && typeof (globalThis as unknown as Record<string, unknown>)[tokens[i]] === 'undefined') {
                        // Coords
                        const position = getIdFromColumnName(tokens[i], true);
                        // Get value
                        let value: unknown;

                        if (typeof obj.options.data[position[1]] != 'undefined' && typeof obj.options.data[position[1]][position[0]] != 'undefined') {
                            value = obj.options.data[position[1]][position[0]];
                        } else {
                            value = '';
                        }
                        // Get column data
                        if (('' + value).substr(0, 1) == '=') {
                            if (typeof formulaResults[tokens[i]] !== 'undefined') {
                                value = formulaResults[tokens[i]];
                            } else {
                                value = execute(value as string, position[0], position[1]);
                                formulaResults[tokens[i]] = value;
                            }
                        }
                        // Type!
                        if (('' + value).trim() == '') {
                            // Null
                            formulaExpressions[tokens[i]] = null;
                        } else {
                            if (value == Number(value) && obj.parent.config.autoCasting != false) {
                                // Number
                                formulaExpressions[tokens[i]] = Number(value);
                            } else {
                                // Trying any formatted number
                                const number = parseNumber.call(obj, value, position[0]);
                                if (obj.parent.config.autoCasting != false && number) {
                                    formulaExpressions[tokens[i]] = number;
                                } else {
                                    // Text as a string literal, with its quotes and backslashes escaped
                                    formulaExpressions[tokens[i]] = JSON.stringify('' + value);
                                }
                            }
                        }
                    }
                }
            }

            let source = text;

            const ret = dispatch.call(obj, 'onbeforeformula', obj, expression, x, y);
            if (ret === false) {
                return expression;
            } else if (ret) {
                expression = ret;
                source = ret;
            }

            // Hand string literals to the formula parser as variables, so characters such as $ or & inside them are not
            // interpreted as part of the formula syntax
            let literalIndex = 0;
            const parsedExpression = expression.substr(1).replace(/"(?:[^"]|"")*"/g, function (literal) {
                const name = '__JSSSTR' + literalIndex++ + '__';
                formulaExpressions[name] = '"' + literal.slice(1, -1).replace(/""/g, '\\"') + '"';
                return name;
            });

            // Convert formula to javascript
            let res;

            try {
                if (formulaEngine) {
                    res = formulaEngine(obj, source, x === undefined ? x : Number(x), y === undefined ? y : Number(y));
                } else {
                    res = formula(parsedExpression, formulaExpressions, x, y, obj);
                }

                if (typeof res === 'function') {
                    res = '#ERROR';
                }
            } catch (e) {
                res = '#ERROR';

                if (obj.parent.config.debugFormulas === true) {
                    console.log(expression.substr(1), formulaExpressions, e);
                }
            }

            return res;
        }
    };

    return execute(expression, x, y);
};

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
        value = executeFormula.call(obj, value as string, i, j);
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
