import jSuites from 'jsuites';
import formula from '@jspreadsheet/formula';

import dispatch from './dispatch.js';
import { refreshSelection, updateCornerPosition } from './selection.js';
import { getColumnName } from './helpers.js';
import { updateMeta } from './meta.js';
import { getFreezeWidth, getIndexWidth } from './freeze.js';
import { updateFreezePosition } from './events.js';
import { updatePagination } from './pagination.js';
import { setFooter } from './footer.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import { setCellLabel, updateGridSize, updateRowIndex, updateTabStop } from './accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import { Worksheet } from '../Worksheet.js';
import type { HistoryCellChange, MergeInfo, Spreadsheet, WorksheetMethod } from '../types/internal';
import type { MaskOptions } from '../types/jsuites';

export const updateTable = function (this: Worksheet) {
    const obj = this;

    // Check for spare
    if (obj.options.minSpareRows! > 0) {
        let numBlankRows = 0;
        for (let j = obj.rows.length - 1; j >= 0; j--) {
            let test = false;
            for (let i = 0; i < obj.headers.length; i++) {
                if (obj.options.data[j][i]) {
                    test = true;
                }
            }
            if (test) {
                break;
            } else {
                numBlankRows++;
            }
        }

        if (obj.options.minSpareRows! - numBlankRows > 0) {
            obj.insertRow(obj.options.minSpareRows! - numBlankRows);
        }
    }

    if (obj.options.minSpareCols! > 0) {
        let numBlankCols = 0;
        for (let i = obj.headers.length - 1; i >= 0; i--) {
            let test = false;
            for (let j = 0; j < obj.rows.length; j++) {
                if (obj.options.data[j][i]) {
                    test = true;
                }
            }
            if (test) {
                break;
            } else {
                numBlankCols++;
            }
        }

        if (obj.options.minSpareCols! - numBlankCols > 0) {
            obj.insertColumn(obj.options.minSpareCols! - numBlankCols);
        }
    }

    // Update footers
    if (obj.options.footers) {
        setFooter.call(obj);
    }

    if (obj.options.columns.length < obj.options.minDimensions[0]) {
        obj.options.minDimensions[0] = obj.options.columns.length;
    }

    // Update corner position
    setTimeout(function () {
        updateCornerPosition.call(obj);
    }, 0);
};

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

// The offset from the column of a cell to the column whose options the cell uses, when it is not 0. See getCellColumnIndex.
// An offset stays valid when a column outside a merged range is inserted, deleted or moved.
const cellColumns = new WeakMap();

/**
 * Render a cell again when getCellColumnIndex gives another column for it.
 * A setting changes only when the options of the two columns are different for it,
 * so a value from setReadOnly or setStyle stays.
 */
export const renderCell = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);

    const element = obj.records[y][x].element;
    const previous = x + (cellColumns.get(element) || 0);
    const c = getCellColumnIndex.call(obj, x, y);
    if (c == previous) {
        return;
    }

    const td = createCell.call(obj, x, y, obj.options.data[y][x], true);

    const getOptions = function (i: number) {
        return (obj.options.columns && obj.options.columns[i]) || {};
    };
    const before = getOptions(previous);
    const after = getOptions(c);

    if ((before.readOnly == true) != (after.readOnly == true)) {
        element.classList.toggle('readonly', after.readOnly == true);
    }
    if (before.align != after.align) {
        element.style.textAlign = td.style.textAlign;
    }
    if (before.wordWrap != after.wordWrap) {
        element.style.whiteSpace = td.style.whiteSpace;
    }
    if (before.type == 'color' || after.type == 'color') {
        element.style.color = td.style.color;
    }
    element.classList.toggle('red', td.classList.contains('red'));
    element.classList.toggle('jss_dropdown', td.classList.contains('jss_dropdown'));

    element.innerHTML = '';
    while (td.firstChild) {
        element.appendChild(td.firstChild);
    }

    if (c == x) {
        cellColumns.delete(element);
    } else {
        cellColumns.set(element, c - x);
    }
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
 * Get dropdown value from key
 */
const getDropDownValue = function (this: Worksheet, column: number, key: unknown) {
    const obj = this;

    const value = [];

    if (obj.options.columns && obj.options.columns[column] && obj.options.columns[column].source) {
        // Create array from source
        const combo = [];
        const source = obj.options.columns[column].source;

        for (let i = 0; i < source.length; i++) {
            if (typeof source[i] == 'object') {
                combo[source[i].id] = source[i].name;
            } else {
                combo[source[i]] = source[i];
            }
        }

        // Guarantee single multiple compatibility
        const delimiter = obj.options.columns[column].delimiter || ';';
        const keys = Array.isArray(key) ? key : ('' + key).split(delimiter);

        for (let i = 0; i < keys.length; i++) {
            if (typeof keys[i] === 'object') {
                value.push(combo[keys[i].id]);
            } else {
                if (combo[keys[i]]) {
                    value.push(combo[keys[i]]);
                }
            }
        }
    } else {
        console.error('Invalid column');
    }

    const displayDelimiter = obj.options.columns[column].delimiter || '; ';
    return value.length > 0 ? value.join(displayDelimiter) : '';
};

// The date, then the date as text, then its parts
const validDate = function (date: jspreadsheet.CellValue | string[] | undefined) {
    date = '' + date;
    if (date.substr(4, 1) == '-' && date.substr(7, 1) == '-') {
        return true;
    } else {
        date = date.split('-');
        // A text without a dash, such as a year, has one part. A part with only digits equals its number.
        if (date[0].length == 4 && (date[0] as unknown) == Number(date[0]) && date[1] && date[1].length == 2 && (date[1] as unknown) == Number(date[1])) {
            return true;
        }
    }
    return false;
};

/**
 * Strip tags. a is the HTML, then the scripts in it. The DOM converts a number, a boolean or undefined to text.
 */
const stripScript = function (a: jspreadsheet.CellValue | undefined | HTMLCollectionOf<HTMLScriptElement>) {
    const b = new Option();
    b.innerHTML = a as string;
    let c = null;
    // A script in b has a parent
    for (a = b.getElementsByTagName('script'); (c = a[0]); ) c.parentNode!.removeChild(c);
    return b.innerHTML;
};

// The DOM converts a number or a boolean value to text, so value is cast to text where the DOM takes text
export const createCell = function (this: Worksheet, i: number, j: number, value: jspreadsheet.CellValue, silent?: boolean) {
    const obj = this;

    // Column options
    const c = getCellColumnIndex.call(obj, i, j);
    const column = obj.options.columns && obj.options.columns[c];

    // Create cell and properties
    const td = document.createElement('td');
    if (c != i) {
        cellColumns.set(td, c - parseInt(i));
    }
    td.setAttribute('data-x', i);
    td.setAttribute('data-y', j);
    td.setAttribute('aria-colindex', i + 2);

    if (obj.headers[i].style.display === 'none') {
        td.style.display = 'none';
    }
    // Security
    if (('' + value).substr(0, 1) == '=' && obj.options.secureFormulas == true) {
        const val = secureFormula(value as string);
        if (val != value) {
            // Update the data container
            value = val;
        }
    }

    // Custom column
    if (column && typeof column.type === 'object') {
        if (obj.parent.config.parseHTML === true) {
            td.innerHTML = value as string;
        } else {
            td.textContent = value as string;
        }
        if (typeof column.type.createCell == 'function') {
            column.type.createCell(td, value, parseInt(i), parseInt(j), obj, column);
        }
    } else {
        // Hidden column
        if (column && column.type == 'hidden') {
            td.style.display = 'none';
            td.textContent = value as string;
        } else if (column && (column.type == 'checkbox' || column.type == 'radio')) {
            // Create input
            const element = document.createElement('input');
            element.type = column.type;
            element.name = 'c' + c;
            element.checked = value == 1 || value == true || value == 'true' ? true : false;
            element.onclick = function () {
                obj.setValue(td, (this as HTMLInputElement).checked);
            };

            if (column.readOnly == true || obj.options.editable == false) {
                element.setAttribute('disabled', 'disabled');
            }

            // Append to the table
            td.appendChild(element);
            // Make sure the values are correct
            obj.options.data[j][i] = element.checked;
        } else if (column && column.type == 'calendar') {
            // Try formatted date
            let formatted = null;
            if (!validDate(value)) {
                const tmp = jSuites.calendar.extractDateFromString(value, (column.options && column.options.format) || 'YYYY-MM-DD');
                if (tmp) {
                    formatted = tmp;
                }
            }
            // Create calendar cell
            // textContent converts a number of hours to text
            td.textContent = jSuites.calendar.getDateString(formatted ? formatted : value, column.options && column.options.format) as string;
        } else if (column && column.type == 'dropdown') {
            // Create dropdown cell
            td.classList.add('jss_dropdown');
            td.textContent = getDropDownValue.call(obj, c, value);
        } else if (column && column.type == 'color') {
            if (column.render == 'square') {
                const color = document.createElement('div');
                color.className = 'color';
                color.style.backgroundColor = value as string;
                td.appendChild(color);
            } else {
                td.style.color = value as string;
                td.textContent = value as string;
            }
        } else if (column && column.type == 'image') {
            // The value can be a number or a boolean
            if (value && ('' + value).substr(0, 10) == 'data:image') {
                const img = document.createElement('img');
                img.src = value as string;
                td.appendChild(img);
            }
        } else {
            if (column && column.type == 'html') {
                td.innerHTML = stripScript(parseValue.call(this, i, j, value, td));
            } else {
                if (obj.parent.config.parseHTML === true) {
                    td.innerHTML = stripScript(parseValue.call(this, i, j, value, td));
                } else {
                    td.textContent = parseValue.call(this, i, j, value, td) as string;
                }
            }
        }
    }

    // Readonly
    if (column && column.readOnly == true) {
        td.className = 'readonly';
    }

    // Text align
    const colAlign = (column && column.align) || obj.options.defaultColAlign || 'center';
    td.style.textAlign = colAlign;

    // Wrap option
    if ((!column || column.wordWrap != false) && (obj.options.wordWrap == true || (column && column.wordWrap == true) || td.innerHTML.length > 200)) {
        td.style.whiteSpace = 'pre-wrap';
    }

    // Overflow
    if (i > 0) {
        if (this.options.textOverflow == true) {
            if (value || td.innerHTML) {
                obj.records[j][i - 1].element.style.overflow = 'hidden';
            } else {
                if (i == obj.options.columns.length - 1) {
                    td.style.overflow = 'hidden';
                }
            }
        }
    }

    // A cell whose content is rendered again keeps its element, so the event is not sent again
    if (!silent) {
        dispatch.call(obj, 'oncreatecell', obj, td, i, j, value);
    }

    return td;
};

/**
 * Update cell content
 *
 * @param {number} x
 * @param {number} y
 * @param {CellValue} value
 * @param {boolean} [force]
 */
// setValue without a value clears the cell, and the data keeps undefined. The DOM converts a number, a boolean or
// undefined to text, so value is cast to text where the DOM takes text.
export const updateCell = function (this: Worksheet, x: number, y: number, value: jspreadsheet.CellValue | undefined, force?: boolean) {
    const obj = this;

    // Column options
    const c = getCellColumnIndex.call(obj, x, y);
    const column = obj.options.columns && obj.options.columns[c];

    let record;

    // Changing value depending on the column type
    if (obj.records[y][x].element.classList.contains('readonly') == true && !force) {
        // Do nothing
        record = {
            x: x,
            y: y,
            col: x,
            row: y,
        };
    } else {
        // Security
        if (('' + value).substr(0, 1) == '=' && obj.options.secureFormulas == true) {
            const val = secureFormula(value as string);
            if (val != value) {
                // Update the data container
                value = val;
            }
        }

        // On change
        const val = dispatch.call(obj, 'onbeforechange', obj, obj.records[y][x].element, x, y, value);

        // If you return something this will overwrite the value
        if (val != undefined) {
            value = val;
        }

        if (column && typeof column.type === 'object' && typeof column.type.updateCell === 'function') {
            const result = column.type.updateCell(obj.records[y][x].element, value, parseInt(x), parseInt(y), obj, column);

            if (result !== undefined) {
                value = result;
            }
        }

        // History format
        record = {
            x: x,
            y: y,
            col: x,
            row: y,
            value: value,
            oldValue: obj.options.data[y][x],
        };

        // A custom editor can have setValue, which is not in the public typings
        const editor = (column && typeof column.type === 'object' ? column.type : null) as
            | (jspreadsheet.CustomEditor & { setValue?: (cell: HTMLTableCellElement, value: jspreadsheet.CellValue | undefined) => void })
            | null;
        if (editor) {
            // Update data and cell
            obj.options.data[y][x] = value as jspreadsheet.CellValue;
            if (typeof editor.setValue === 'function') {
                editor.setValue(obj.records[y][x].element, value);
            }
        } else {
            // Native functions
            if (column && (column.type == 'checkbox' || column.type == 'radio')) {
                // Unchecked all options
                if (column.type == 'radio') {
                    for (let j = 0; j < obj.options.data.length; j++) {
                        obj.options.data[j][c] = false;
                    }
                }

                // Update data and cell
                (obj.records[y][x].element.children[0] as HTMLInputElement).checked =
                    value == 1 || value == true || value == 'true' || value == 'TRUE' ? true : false;
                obj.options.data[y][x] = (obj.records[y][x].element.children[0] as HTMLInputElement).checked;
            } else if (column && column.type == 'dropdown') {
                // Update data and cell
                obj.options.data[y][x] = value as jspreadsheet.CellValue;
                obj.records[y][x].element.textContent = getDropDownValue.call(obj, c, value);
            } else if (column && column.type == 'calendar') {
                // Try formatted date
                let formatted = null;
                if (!validDate(value)) {
                    const tmp = jSuites.calendar.extractDateFromString(value, (column.options && column.options.format) || 'YYYY-MM-DD');
                    if (tmp) {
                        formatted = tmp;
                    }
                }
                // Update data and cell
                obj.options.data[y][x] = value as jspreadsheet.CellValue;
                // textContent converts a number of hours to text
                obj.records[y][x].element.textContent = jSuites.calendar.getDateString(
                    formatted ? formatted : value,
                    column.options && column.options.format
                ) as string;
            } else if (column && column.type == 'color') {
                // Update color
                obj.options.data[y][x] = value as jspreadsheet.CellValue;
                // Render
                if (column.render == 'square') {
                    const color = document.createElement('div');
                    color.className = 'color';
                    color.style.backgroundColor = value as string;
                    obj.records[y][x].element.textContent = '';
                    obj.records[y][x].element.appendChild(color);
                } else {
                    obj.records[y][x].element.style.color = value as string;
                    obj.records[y][x].element.textContent = value as string;
                }
            } else if (column && column.type == 'image') {
                value = '' + value;
                obj.options.data[y][x] = value;
                obj.records[y][x].element.innerHTML = '';
                if (value && value.substr(0, 10) == 'data:image') {
                    const img = document.createElement('img');
                    img.src = value;
                    obj.records[y][x].element.appendChild(img);
                }
            } else {
                // Update data and cell
                obj.options.data[y][x] = value as jspreadsheet.CellValue;
                // Label
                if (column && column.type == 'html') {
                    obj.records[y][x].element.innerHTML = stripScript(parseValue.call(obj, x, y, value));
                } else {
                    if (obj.parent.config.parseHTML === true) {
                        obj.records[y][x].element.innerHTML = stripScript(parseValue.call(obj, x, y, value, obj.records[y][x].element));
                    } else {
                        obj.records[y][x].element.textContent = parseValue.call(obj, x, y, value, obj.records[y][x].element) as string;
                    }
                }
                // Handle big text inside a cell
                if (
                    (!column || column.wordWrap != false) &&
                    (obj.options.wordWrap == true || (column && column.wordWrap == true) || obj.records[y][x].element.innerHTML.length > 200)
                ) {
                    obj.records[y][x].element.style.whiteSpace = 'pre-wrap';
                } else {
                    obj.records[y][x].element.style.whiteSpace = '';
                }
            }
        }

        // Overflow
        if (x > 0) {
            if (value) {
                obj.records[y][x - 1].element.style.overflow = 'hidden';
            } else {
                obj.records[y][x - 1].element.style.overflow = '';
            }
        }

        if (column && typeof column.render === 'function') {
            // updateCell has read obj.records[y][x] already, so the cell is there
            column.render(
                (obj.records[y] && obj.records[y][x] ? obj.records[y][x].element : null) as HTMLTableCellElement,
                value,
                parseInt(x),
                parseInt(y),
                obj,
                column
            );
        }

        // Accessible name of the active cell
        if (obj.selection.activeCell && obj.records[y] && obj.records[y][x] && obj.records[y][x].element === obj.selection.activeCell) {
            setCellLabel.call(obj, obj.selection.activeCell);
        }

        // On change. updateCell has read obj.records[y][x] already, so the cell is there.
        dispatch.call(
            obj,
            'onchange',
            obj,
            (obj.records[y] && obj.records[y][x] ? obj.records[y][x].element : null) as HTMLTableCellElement,
            x,
            y,
            value,
            record.oldValue
        );
    }

    return record;
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

/**
 * Secure formula
 */
const secureFormula = function (oldValue: string) {
    let newValue = '';
    let inside = 0;

    for (let i = 0; i < oldValue.length; i++) {
        if (oldValue[i] == '"') {
            if (inside == 0) {
                inside = 1;
            } else {
                inside = 0;
            }
        }

        if (inside == 1) {
            newValue += oldValue[i];
        } else {
            newValue += oldValue[i].toUpperCase();
        }
    }

    return newValue;
};

/**
 * Update all related cells in the chain
 */
// The cells in progress, by cell name
let chainLoopProtection: Record<string, boolean> = {};

export const updateFormulaChain = function (this: Worksheet, x: number, y: number, records: HistoryCellChange[]) {
    const obj = this;

    const cellId = getColumnNameFromId([x, y]);
    if (obj.formulas.chain[cellId] && obj.formulas.chain[cellId].length > 0) {
        if (chainLoopProtection[cellId]) {
            obj.records[y][x].element.innerHTML = '#ERROR';
            // The chain ends here
            obj.formulas.chain[cellId] = [];
        } else {
            // Protection
            chainLoopProtection[cellId] = true;

            for (let i = 0; i < obj.formulas.chain[cellId].length; i++) {
                const cell = getIdFromColumnName(obj.formulas.chain[cellId][i], true);
                // Update cell
                const value = '' + obj.options.data[cell[1]][cell[0]];
                if (value.substr(0, 1) == '=') {
                    records.push(updateCell.call(obj, cell[0], cell[1], value, true));
                } else {
                    // No longer a formula, remove from the chain
                    // This changes nothing: Object.keys returns a new array (issue #43)
                    Object.keys(obj.formulas.chain)[i] = null!;
                }
                updateFormulaChain.call(obj, cell[0], cell[1], records);
            }
        }
    }

    chainLoopProtection = {};
};

/**
 * Update formula
 */
export const updateFormula = function (formula: string, referencesToUpdate: Record<string, string>) {
    const testLetter = /[A-Z]/;
    const testNumber = /[0-9]/;

    let newFormula = '';
    let letter = null;
    let number = null;
    let token = '';

    for (let index = 0; index < formula.length; index++) {
        if (testLetter.exec(formula[index])) {
            letter = 1;
            number = 0;
            token += formula[index];
        } else if (testNumber.exec(formula[index])) {
            number = letter ? 1 : 0;
            token += formula[index];
        } else {
            if (letter && number) {
                token = referencesToUpdate[token] ? referencesToUpdate[token] : token;
            }
            newFormula += token;
            newFormula += formula[index];
            letter = 0;
            number = 0;
            token = '';
        }
    }

    if (token) {
        if (letter && number) {
            token = referencesToUpdate[token] ? referencesToUpdate[token] : token;
        }
        newFormula += token;
    }

    return newFormula;
};

/**
 * Update formulas
 */
const updateFormulas = function (this: Worksheet, referencesToUpdate: Record<string, string>) {
    const obj = this;

    // Update formulas
    for (let j = 0; j < obj.options.data.length; j++) {
        for (let i = 0; i < obj.options.data[0].length; i++) {
            const value = '' + obj.options.data[j][i];
            // Is formula
            if (value.substr(0, 1) == '=') {
                // Replace tokens
                const newFormula = updateFormula(value, referencesToUpdate);
                if (newFormula != value) {
                    obj.options.data[j][i] = newFormula;
                }
            }
        }
    }

    // Update formula chain
    obj.formulas.rename(referencesToUpdate);
};

/**
 * Update cell references
 *
 * @return void
 */
export const updateTableReferences = function (this: Worksheet) {
    const obj = this;
    if (obj.skipUpdateTableReferences) {
        return;
    }

    // Update headers
    for (let i = 0; i < obj.headers.length; i++) {
        const x = parseInt(obj.headers[i].getAttribute('data-x'));

        if (x != i) {
            // Update coords
            obj.headers[i].setAttribute('data-x', i);
            obj.headers[i].setAttribute('aria-colindex', i + 2);
            // Title
            if (!obj.headers[i].getAttribute('title')) {
                obj.headers[i].innerHTML = getColumnName(i);
            }
        }
    }

    // Update all rows
    for (let j = 0; j < obj.rows.length; j++) {
        if (obj.rows[j]) {
            const y = parseInt(obj.rows[j].element.getAttribute('data-y'));

            if (y != j) {
                // Update coords
                obj.rows[j].element.setAttribute('data-y', j);
                obj.rows[j].element.children[0].setAttribute('data-y', j);
                updateRowIndex.call(obj, obj.rows[j].element, j);
                // Row number
                obj.rows[j].element.children[0].innerHTML = String(j + 1);
            }
        }
    }

    // Regular cells affected by this change: the new cell name by the old one
    const affectedTokens: Record<string, string> = {};
    // The merged cells affected by this change: false, or the new cell name and the column and row offsets
    const mergeCellUpdates: Record<string, false | [string, number, number]> = {};

    // Update cell
    const updatePosition = function (x: number, y: number, i: number, j: number) {
        if (x != i) {
            obj.records[j][i].element.setAttribute('data-x', i);
            obj.records[j][i].element.setAttribute('aria-colindex', i + 2);
        }
        if (y != j) {
            obj.records[j][i].element.setAttribute('data-y', j);
        }

        // Other updates
        if (x != i || y != j) {
            const columnIdFrom = getColumnNameFromId([x, y]);
            const columnIdTo = getColumnNameFromId([i, j]);
            affectedTokens[columnIdFrom] = columnIdTo;
        }
    };

    for (let j = 0; j < obj.records.length; j++) {
        for (let i = 0; i < obj.records[0].length; i++) {
            if (obj.records[j][i]) {
                // Current values
                const x = parseInt(obj.records[j][i].element.getAttribute('data-x'));
                const y = parseInt(obj.records[j][i].element.getAttribute('data-y'));

                // Update column
                if (obj.records[j][i].element.getAttribute('data-merged')) {
                    const columnIdFrom = getColumnNameFromId([x, y]);
                    const columnIdTo = getColumnNameFromId([i, j]);
                    if (mergeCellUpdates[columnIdFrom] == null) {
                        if (columnIdFrom == columnIdTo) {
                            mergeCellUpdates[columnIdFrom] = false;
                        } else {
                            const totalX = i - x;
                            const totalY = j - y;
                            mergeCellUpdates[columnIdFrom] = [columnIdTo, totalX, totalY];
                        }
                    }
                } else {
                    updatePosition(x, y, i, j);
                }
            }
        }
    }

    // Update merged if applicable
    const keys = Object.keys(mergeCellUpdates);
    if (keys.length) {
        for (let i = 0; i < keys.length; i++) {
            const update = mergeCellUpdates[keys[i]];
            if (update) {
                const info = getIdFromColumnName(keys[i], true);
                let x = info[0];
                let y = info[1];
                updatePosition(x, y, x + update[1], y + update[2]);

                const columnIdFrom = keys[i];
                const columnIdTo = update[0];
                for (let j = 0; j < (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]!.length; j++) {
                    x = parseInt((obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-x'));
                    y = parseInt((obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-y'));
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-x', x + update[1]);
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('aria-colindex', x + update[1] + 2);
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-y', y + update[2]);
                }

                obj.options.mergeCells![columnIdTo] = obj.options.mergeCells![columnIdFrom];
                delete obj.options.mergeCells![columnIdFrom];
            }
        }
    }

    // Update formulas
    updateFormulas.call(obj, affectedTokens);

    // Update meta data
    updateMeta.call(obj, affectedTokens);

    // Update the size of the grid
    updateGridSize.call(obj);

    // The tab stop of the grid can be in a deleted row or column
    updateTabStop.call(obj);

    // Refresh selection
    refreshSelection.call(obj);

    // Update table with custom configuration if applicable
    updateTable.call(obj);
};

/**
 * Update scroll position based on the selection
 */
export const updateScroll = function (this: Worksheet, direction: 0 | 1 | 2 | 3) {
    const obj = this;

    // Jspreadsheet Container information
    const contentRect = obj.content.getBoundingClientRect();
    const x1 = contentRect.left;
    const y1 = contentRect.top;
    const w1 = contentRect.width;
    const h1 = contentRect.height;

    // Direction Left or Up
    const reference = obj.records[obj.selection.range![3]][obj.selection.range![2]].element;

    // Reference
    const referenceRect = reference.getBoundingClientRect();
    const x2 = referenceRect.left;
    const y2 = referenceRect.top;
    const w2 = referenceRect.width;
    const h2 = referenceRect.height;

    let x, y;

    // Direction
    if (direction == 0 || direction == 1) {
        x = x2 - x1 + obj.content.scrollLeft;
        y = y2 - y1 + obj.content.scrollTop - 2;
    } else {
        x = x2 - x1 + obj.content.scrollLeft + w2;
        y = y2 - y1 + obj.content.scrollTop + h2;
    }

    // Top position check
    if (y > obj.content.scrollTop + 30 && y < obj.content.scrollTop + h1) {
        // In the viewport
    } else {
        // Out of viewport
        if (y < obj.content.scrollTop + 30) {
            obj.content.scrollTop = y - h2;
        } else {
            obj.content.scrollTop = y - (h1 - 2);
        }
    }

    // Freeze columns?
    const freezed = getFreezeWidth.call(obj);

    // The frozen columns stay at the left of the content only when scrollLeft is more than the width of the index column.
    // Then a cell whose left border is under the frozen columns goes to the right of the frozen columns.
    const indexWidth = getIndexWidth.call(obj);
    const frozen = freezed > 0 && obj.content.scrollLeft > indexWidth;
    const left = x2 - x1 + obj.content.scrollLeft;
    const underFreezed = frozen && left < obj.content.scrollLeft + freezed;

    // Left position check - TODO: change that to the bottom border of the element
    if (
        (frozen && obj.selection.range![2] < obj.options.freezeColumns!) ||
        (!underFreezed && x > obj.content.scrollLeft + freezed && x < obj.content.scrollLeft + w1)
    ) {
        // In the viewport
    } else if (underFreezed) {
        obj.content.scrollLeft = left - freezed - 1 < indexWidth ? 0 : left - freezed - 1;
    } else {
        // Out of viewport
        if (x < obj.content.scrollLeft + 30) {
            obj.content.scrollLeft = x;
            if (obj.content.scrollLeft < 50) {
                obj.content.scrollLeft = 0;
            }
        } else if (x < obj.content.scrollLeft + freezed) {
            obj.content.scrollLeft = x - freezed - 1;
        } else {
            obj.content.scrollLeft = x - (w1 - 20);
        }
    }
};

export const updateResult = function (this: Worksheet) {
    const obj = this;

    let total = 0;
    let index = 0;

    // Page 1
    if (obj.options.lazyLoading == true) {
        total = 100;
    } else if (obj.options.pagination! > 0) {
        total = obj.options.pagination!;
    } else {
        if (obj.results) {
            total = obj.results.length;
        } else {
            total = obj.rows.length;
        }
    }

    // Reset current nodes
    while (obj.tbody.firstChild) {
        obj.tbody.removeChild(obj.tbody.firstChild);
    }

    // Hide all records from the table
    for (let j = 0; j < obj.rows.length; j++) {
        if (!obj.results || obj.results.indexOf(j) > -1) {
            if (index < total) {
                obj.tbody.appendChild(obj.rows[j].element);
                index++;
            }
            obj.rows[j].element.style.display = '';
        } else {
            obj.rows[j].element.style.display = 'none';
        }
    }

    // The tab stop of the grid can be in a row that is not in the results
    updateTabStop.call(obj);

    // Update pagination
    if (obj.options.pagination! > 0) {
        updatePagination.call(obj);
    }

    updateCornerPosition.call(obj);

    dispatch.call(obj, 'onupdateresult', obj, obj.results);

    return total;
};

/**
 * Get the cell object
 *
 * @param object cell
 * @return string value
 */
export const getCell = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    if (typeof x === 'string') {
        // Convert in case name is excel liked ex. A10, BB92
        const cell = getIdFromColumnName(x, true);

        x = cell[0];
        y = cell[1];
    }

    return obj.records[y][x].element;
};

/**
 * Get the cell object from coords
 *
 * @param object cell
 * @return string value
 */
export const getCellFromCoords: WorksheetMethod<'getCellFromCoords'> = function (x, y) {
    const obj = this;

    return obj.records[y][x].element;
};

/**
 * Get label
 *
 * @param object cell
 * @return string value
 */
export const getLabel = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    if (typeof x === 'string') {
        // Convert in case name is excel liked ex. A10, BB92
        const cell = getIdFromColumnName(x, true);

        x = cell[0];
        y = cell[1];
    }

    return obj.records[y][x].element.innerHTML;
};

/**
 * Activate/Disable fullscreen
 * use programmatically : table.fullscreen(); or table.fullscreen(true); or table.fullscreen(false);
 * @Param {boolean} activate
 */
export const fullscreen = function (this: Spreadsheet, activate?: boolean) {
    const spreadsheet = this;

    // If activate not defined, get reverse options.fullscreen
    if (activate == null) {
        activate = !spreadsheet.config.fullscreen;
    }

    // If change
    if (spreadsheet.config.fullscreen != activate) {
        spreadsheet.config.fullscreen = activate;

        // Test LazyLoading conflict
        if (activate == true) {
            spreadsheet.element.classList.add('fullscreen');
        } else {
            spreadsheet.element.classList.remove('fullscreen');
        }
    }
};

/**
 * Show index column
 */
export const showIndex: WorksheetMethod<'showIndex'> = function () {
    const obj = this;

    obj.table.classList.remove('jss_hidden_index');

    // The position of the frozen columns depends on the width of the index column
    if (obj.options.freezeColumns! > 0) {
        updateFreezePosition.call(obj);
    }
};

/**
 * Hide index column
 */
export const hideIndex: WorksheetMethod<'hideIndex'> = function () {
    const obj = this;

    obj.table.classList.add('jss_hidden_index');

    // The position of the frozen columns depends on the width of the index column
    if (obj.options.freezeColumns! > 0) {
        updateFreezePosition.call(obj);
    }
};

/**
 * Create a nested header object
 */
// The row of nested headers keeps its element
export const createNestedHeader = function (this: Worksheet, nestedInformation: jspreadsheet.NestedHeaderCell[] & { element?: HTMLTableRowElement }) {
    const obj = this;

    const tr = document.createElement('tr');
    tr.classList.add('jss_nested');
    const td = document.createElement('td');
    td.classList.add('jss_selectall');

    tr.appendChild(td);
    // Element
    nestedInformation.element = tr;

    let headerIndex = 0;
    for (let i = 0; i < nestedInformation.length; i++) {
        // Default values
        if (!nestedInformation[i].colspan) {
            nestedInformation[i].colspan = 1;
        }
        if (!nestedInformation[i].title) {
            nestedInformation[i].title = '';
        }
        if (!nestedInformation[i].id) {
            nestedInformation[i].id = '';
        }

        // Number of columns. colspan, title and id have default values above.
        let numberOfColumns = nestedInformation[i].colspan!;

        // Classes container
        const column = [];
        // Header classes for this cell
        for (let x = 0; x < numberOfColumns; x++) {
            if (obj.options.columns[headerIndex] && obj.options.columns[headerIndex].type == 'hidden') {
                numberOfColumns++;
            }
            column.push(headerIndex);
            headerIndex++;
        }

        // Created the nested cell
        const td = document.createElement('td');
        td.setAttribute('data-column', column.join(','));
        td.setAttribute('colspan', nestedInformation[i].colspan!);
        td.setAttribute('align', nestedInformation[i].align || 'center');
        td.setAttribute('id', nestedInformation[i].id!);
        td.textContent = nestedInformation[i].title!;
        tr.appendChild(td);
    }

    return tr;
};

// Spreadsheets call it too
export const getWorksheetActive = function (this: Worksheet | Spreadsheet): number {
    const spreadsheet = this instanceof Worksheet ? this.parent : this;

    // The element of a spreadsheet with tabs has the tabs object
    // getActive returns false only when no tab is open, and a tab is open while the spreadsheet has worksheets
    return spreadsheet.element.tabs ? (spreadsheet.element.tabs.getActive() as number) : 0;
};

export const getWorksheetInstance = function (this: Spreadsheet, index?: number) {
    const spreadsheet = this;

    const worksheetIndex = typeof index !== 'undefined' ? index : getWorksheetActive.call(spreadsheet);

    return spreadsheet.worksheets[worksheetIndex];
};
