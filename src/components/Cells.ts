import { getCoordsFromCellName } from '../utils/helpers.js';
import { getIdFromColumnName, getColumnNameFromId } from '../utils/internalHelpers.js';
import jSuites from 'jsuites';
import { getMask, isFormula } from '../utils/values.js';
import type { Worksheet } from '../types/internal';
import type jspreadsheet from '../jspreadsheet';

// The date, then the date as text, then its parts
function validDate(date: jspreadsheet.CellValue | string[] | undefined) {
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
}

/**
 * Strip tags. a is the HTML, then the scripts in it. The DOM converts a number, a boolean or undefined to text.
 */
function stripScript(a: jspreadsheet.CellValue | undefined | HTMLCollectionOf<HTMLScriptElement>) {
    const b = new Option();
    b.innerHTML = a as string;
    let c = null;
    // A script in b has a parent
    for (a = b.getElementsByTagName('script'); (c = a[0]); ) c.parentNode!.removeChild(c);
    return b.innerHTML;
}

/**
 * Secure formula
 */
function secureFormula(oldValue: string) {
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
}

// The offset from the column of a cell to the column whose options the cell uses, when it is not 0. See getColumnIndex.
// An offset stays valid when a column outside a merged range is inserted, deleted or moved.
const cellColumns = new WeakMap();

/**
 * The cells of a worksheet: creating and rendering them, the value that a cell shows, finding the td of a cell and its
 * label, and whether it is read-only.
 *
 * Each worksheet has one, as worksheet.cells. getCell, getCellFromCoords, getLabel, setReadOnly and isReadOnly call it.
 */
export class Cells {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Make a cell read-only, or editable again.
     * @param cell - The cell name, for example "B3", or the td of the cell.
     * @param state - Whether the cell is read-only.
     */
    setReadOnly(cell: string | HTMLElement, state: boolean) {
        const worksheet = this.worksheet;

        let record;

        if (typeof cell === 'string') {
            const coords = getCoordsFromCellName(cell);

            record = worksheet.records[coords![1]!][coords![0]];
        } else {
            const x = parseInt(cell.getAttribute('data-x'));
            const y = parseInt(cell.getAttribute('data-y'));

            record = worksheet.records[y][x];
        }

        if (state) {
            record.element.classList.add('readonly');
        } else {
            record.element.classList.remove('readonly');
        }
    }

    /**
     * Check if a cell is read-only.
     * @param x - The column, or the cell name, for example "B3".
     * @param y - The row.
     */
    isReadOnly(x: number | string, y?: number): boolean {
        const worksheet = this.worksheet;

        if (typeof x === 'string' && typeof y === 'undefined') {
            const coords = getCoordsFromCellName(x);

            [x, y] = coords as [number, number];
        }

        return worksheet.records[y!][x as number].element.classList.contains('readonly');
    }

    /**
     * Get the cell object
     *
     * @param object cell
     * @return string value
     */
    // A cell name, or the coordinates of the cell
    get(x: number | string, y?: number) {
        const worksheet = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x, true);

            x = cell[0];
            y = cell[1];
        }

        // y is set with the coordinates, and from the name
        return worksheet.records[y!][x].element;
    }

    /**
     * Get the cell object from coords
     *
     * @param object cell
     * @return string value
     */
    getFromCoords(x: number, y: number) {
        const worksheet = this.worksheet;

        return worksheet.records[y][x].element;
    }

    /**
     * Get label
     *
     * @param object cell
     * @return string value
     */
    // A cell name, or the coordinates of the cell
    getLabel(x: number | string, y?: number) {
        const worksheet = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x, true);

            x = cell[0];
            y = cell[1];
        }

        // y is set with the coordinates, and from the name
        return worksheet.records[y!][x].element.innerHTML;
    }

    /**
     * Render a cell again when getColumnIndex gives another column for it.
     * A setting changes only when the options of the two columns are different for it,
     * so a value from setReadOnly or setStyle stays.
     */
    render(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);

        const element = worksheet.records[y][x].element;
        const previous = x + (cellColumns.get(element) || 0);
        const c = this.getColumnIndex(x, y);
        if (c == previous) {
            return;
        }

        const td = this.create(x, y, worksheet.options.data[y][x], true);

        function getOptions(i: number) {
            return (worksheet.options.columns && worksheet.options.columns[i]) || {};
        }
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
    }

    /**
     * Get dropdown value from key
     */
    private getDropDownValue(column: number, key: unknown) {
        const worksheet = this.worksheet;

        const value = [];

        if (worksheet.options.columns && worksheet.options.columns[column] && worksheet.options.columns[column].source) {
            // Create array from source
            const combo = [];
            const source = worksheet.options.columns[column].source;

            for (let i = 0; i < source.length; i++) {
                if (typeof source[i] == 'object') {
                    combo[source[i].id] = source[i].name;
                } else {
                    combo[source[i]] = source[i];
                }
            }

            // Guarantee single multiple compatibility
            const delimiter = worksheet.options.columns[column].delimiter || ';';
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

        const displayDelimiter = worksheet.options.columns[column].delimiter || '; ';
        return value.length > 0 ? value.join(displayDelimiter) : '';
    }

    // The DOM converts a number or a boolean value to text, so value is cast to text where the DOM takes text
    create(i: number, j: number, value: jspreadsheet.CellValue, silent?: boolean) {
        const worksheet = this.worksheet;

        // Column options
        const c = this.getColumnIndex(i, j);
        const column = worksheet.options.columns && worksheet.options.columns[c];

        // Create cell and properties
        const td = document.createElement('td');
        if (c != i) {
            cellColumns.set(td, c - parseInt(i));
        }
        td.setAttribute('data-x', i);
        td.setAttribute('data-y', j);
        td.setAttribute('aria-colindex', i + 2);

        if (worksheet.headers[i].style.display === 'none') {
            td.style.display = 'none';
        }
        // Security
        if (('' + value).substr(0, 1) == '=' && worksheet.options.secureFormulas == true) {
            const val = secureFormula(value as string);
            if (val != value) {
                // Update the data container
                value = val;
            }
        }

        // Custom column
        if (column && typeof column.type === 'object') {
            if (worksheet.parent.config.parseHTML === true) {
                td.innerHTML = value as string;
            } else {
                td.textContent = value as string;
            }
            if (typeof column.type.createCell == 'function') {
                column.type.createCell(td, value, parseInt(i), parseInt(j), worksheet, column);
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
                element.onclick = () => {
                    worksheet.setValue(td, element.checked);
                };

                if (column.readOnly == true || worksheet.options.editable == false) {
                    element.setAttribute('disabled', 'disabled');
                }

                // Append to the table
                td.appendChild(element);
                // Make sure the values are correct
                worksheet.options.data[j][i] = element.checked;
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
                td.textContent = this.getDropDownValue(c, value);
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
                    td.innerHTML = stripScript(this.parseValue(i, j, value, td));
                } else {
                    if (worksheet.parent.config.parseHTML === true) {
                        td.innerHTML = stripScript(this.parseValue(i, j, value, td));
                    } else {
                        td.textContent = this.parseValue(i, j, value, td) as string;
                    }
                }
            }
        }

        // Readonly
        if (column && column.readOnly == true) {
            td.className = 'readonly';
        }

        // Text align
        const colAlign = (column && column.align) || worksheet.options.defaultColAlign || 'center';
        td.style.textAlign = colAlign;

        // Wrap option
        if ((!column || column.wordWrap != false) && (worksheet.options.wordWrap == true || (column && column.wordWrap == true) || td.innerHTML.length > 200)) {
            td.style.whiteSpace = 'pre-wrap';
        }

        // Overflow
        if (i > 0) {
            if (worksheet.options.textOverflow == true) {
                if (value || td.innerHTML) {
                    worksheet.records[j][i - 1].element.style.overflow = 'hidden';
                } else {
                    if (i == worksheet.options.columns.length - 1) {
                        td.style.overflow = 'hidden';
                    }
                }
            }
        }

        // A cell whose content is rendered again keeps its element, so the event is not sent again
        if (!silent) {
            worksheet.dispatch('oncreatecell', worksheet, td, i, j, value);
        }

        return td;
    }

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
    update(x: number, y: number, value: jspreadsheet.CellValue | undefined, force?: boolean) {
        const worksheet = this.worksheet;

        // Column options
        const c = this.getColumnIndex(x, y);
        const column = worksheet.options.columns && worksheet.options.columns[c];

        let record;

        // Changing value depending on the column type
        if (worksheet.records[y][x].element.classList.contains('readonly') == true && !force) {
            // Do nothing
            record = {
                x,
                y,
                col: x,
                row: y,
            };
        } else {
            // Security
            if (('' + value).substr(0, 1) == '=' && worksheet.options.secureFormulas == true) {
                const val = secureFormula(value as string);
                if (val != value) {
                    // Update the data container
                    value = val;
                }
            }

            // On change
            const val = worksheet.dispatch('onbeforechange', worksheet, worksheet.records[y][x].element, x, y, value);

            // If you return something this will overwrite the value
            if (val != undefined) {
                value = val;
            }

            if (column && typeof column.type === 'object' && typeof column.type.updateCell === 'function') {
                const result = column.type.updateCell(worksheet.records[y][x].element, value, parseInt(x), parseInt(y), worksheet, column);

                if (result !== undefined) {
                    value = result;
                }
            }

            // History format
            record = {
                x,
                y,
                col: x,
                row: y,
                value,
                oldValue: worksheet.options.data[y][x],
            };

            // A custom editor can have setValue, which is not in the public typings
            const editor = (column && typeof column.type === 'object' ? column.type : null) as
                | (jspreadsheet.CustomEditor & { setValue?: (cell: HTMLTableCellElement, value: jspreadsheet.CellValue | undefined) => void })
                | null;
            if (editor) {
                // Update data and cell
                worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                if (typeof editor.setValue === 'function') {
                    editor.setValue(worksheet.records[y][x].element, value);
                }
            } else {
                // Native functions
                if (column && (column.type == 'checkbox' || column.type == 'radio')) {
                    // Unchecked all options
                    if (column.type == 'radio') {
                        for (let j = 0; j < worksheet.options.data.length; j++) {
                            worksheet.options.data[j][c] = false;
                        }
                    }

                    // Update data and cell
                    (worksheet.records[y][x].element.children[0] as HTMLInputElement).checked =
                        value == 1 || value == true || value == 'true' || value == 'TRUE' ? true : false;
                    worksheet.options.data[y][x] = (worksheet.records[y][x].element.children[0] as HTMLInputElement).checked;
                } else if (column && column.type == 'dropdown') {
                    // Update data and cell
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                    worksheet.records[y][x].element.textContent = this.getDropDownValue(c, value);
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
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                    // textContent converts a number of hours to text
                    worksheet.records[y][x].element.textContent = jSuites.calendar.getDateString(
                        formatted ? formatted : value,
                        column.options && column.options.format
                    ) as string;
                } else if (column && column.type == 'color') {
                    // Update color
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                    // Render
                    if (column.render == 'square') {
                        const color = document.createElement('div');
                        color.className = 'color';
                        color.style.backgroundColor = value as string;
                        worksheet.records[y][x].element.textContent = '';
                        worksheet.records[y][x].element.appendChild(color);
                    } else {
                        worksheet.records[y][x].element.style.color = value as string;
                        worksheet.records[y][x].element.textContent = value as string;
                    }
                } else if (column && column.type == 'image') {
                    value = '' + value;
                    worksheet.options.data[y][x] = value;
                    worksheet.records[y][x].element.innerHTML = '';
                    if (value && value.substr(0, 10) == 'data:image') {
                        const img = document.createElement('img');
                        img.src = value;
                        worksheet.records[y][x].element.appendChild(img);
                    }
                } else {
                    // Update data and cell
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                    // Label
                    if (column && column.type == 'html') {
                        worksheet.records[y][x].element.innerHTML = stripScript(this.parseValue(x, y, value));
                    } else {
                        if (worksheet.parent.config.parseHTML === true) {
                            worksheet.records[y][x].element.innerHTML = stripScript(this.parseValue(x, y, value, worksheet.records[y][x].element));
                        } else {
                            worksheet.records[y][x].element.textContent = this.parseValue(x, y, value, worksheet.records[y][x].element) as string;
                        }
                    }
                    // Handle big text inside a cell
                    if (
                        (!column || column.wordWrap != false) &&
                        (worksheet.options.wordWrap == true || (column && column.wordWrap == true) || worksheet.records[y][x].element.innerHTML.length > 200)
                    ) {
                        worksheet.records[y][x].element.style.whiteSpace = 'pre-wrap';
                    } else {
                        worksheet.records[y][x].element.style.whiteSpace = '';
                    }
                }
            }

            // Overflow
            if (x > 0) {
                if (value) {
                    worksheet.records[y][x - 1].element.style.overflow = 'hidden';
                } else {
                    worksheet.records[y][x - 1].element.style.overflow = '';
                }
            }

            if (column && typeof column.render === 'function') {
                // updateCell has read worksheet.records[y][x] already, so the cell is there
                column.render(
                    (worksheet.records[y] && worksheet.records[y][x] ? worksheet.records[y][x].element : null) as HTMLTableCellElement,
                    value,
                    parseInt(x),
                    parseInt(y),
                    worksheet,
                    column
                );
            }

            // Accessible name of the active cell
            if (
                worksheet.selection.activeCell &&
                worksheet.records[y] &&
                worksheet.records[y][x] &&
                worksheet.records[y][x].element === worksheet.selection.activeCell
            ) {
                worksheet.accessibility.setCellLabel(worksheet.selection.activeCell);
            }

            // On change. updateCell has read worksheet.records[y][x] already, so the cell is there.
            worksheet.dispatch(
                'onchange',
                worksheet,
                (worksheet.records[y] && worksheet.records[y][x] ? worksheet.records[y][x].element : null) as HTMLTableCellElement,
                x,
                y,
                value,
                record.oldValue
            );
        }

        return record;
    }

    /**
     * Get the column whose options a cell uses. The first cell of a merged range whose first column has the
     * type hidden uses the options of the first visible column of the range, because it shows the range.
     * A checkbox or radio column is not used, because these types write a boolean to the data of the cell.
     */
    getColumnIndex(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);

        if (worksheet.options.columns && worksheet.options.columns[x] && worksheet.options.columns[x].type == 'hidden' && worksheet.options.mergeCells) {
            const merged = worksheet.options.mergeCells[getColumnNameFromId([x, y])];
            if (merged) {
                for (let i = x + 1; i < x + merged[0]; i++) {
                    if (
                        worksheet.headers[i] &&
                        worksheet.headers[i].style.display != 'none' &&
                        (!worksheet.options.columns[i] || worksheet.options.columns[i].type != 'hidden')
                    ) {
                        if (worksheet.options.columns[i] && (worksheet.options.columns[i].type == 'checkbox' || worksheet.options.columns[i].type == 'radio')) {
                            break;
                        }
                        return i;
                    }
                }
            }
        }

        return x;
    }

    /**
     * The value that a cell shows: the result of a formula, or the value with the mask of the column. The DOM converts a
     * number, a boolean or undefined to text.
     */
    parseValue(i: number, j: number, value: jspreadsheet.CellValue | undefined, cell?: HTMLElement) {
        const worksheet = this.worksheet;

        if (('' + value).substr(0, 1) == '=' && worksheet.parent.config.parseFormulas != false) {
            // The cell shows the result as text, as the DOM converts it. An array from a range shows its items with commas.
            value = worksheet.formulas.execute(value as string, i, j) as jspreadsheet.CellValue;
        }

        // Column options
        const options = worksheet.options.columns && worksheet.options.columns[this.getColumnIndex(i, j)];
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
    }
}
