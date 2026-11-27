import { getCoordsFromCellName } from '../utils/helpers.js';
import { getIdFromColumnName, getColumnNameFromId } from '../utils/internalHelpers.js';
import { getMask, isFormula } from '../utils/values.js';
import { formatDate, parseDate } from '../utils/dates.js';
import { render } from '../utils/mask.js';
import type { HistoryCellChange, Worksheet } from '../types/internal';
import type jspreadsheet from '../jspreadsheet';

// The date, then the date as text, then its parts
function validDate(date: jspreadsheet.CellValue | string[] | undefined) {
    date = `${date}`;
    if (date.substr(4, 1) === '-' && date.substr(7, 1) === '-') {
        return true;
    } else {
        date = date.split('-');
        // A text without a dash, such as a year, has one part. A part with only digits equals its number.
        // eslint-disable-next-line eqeqeq -- a text with only digits equals its number
        if (date[0].length === 4 && (date[0] as unknown) == Number(date[0]) && date[1]?.length === 2 && (date[1] as unknown) == Number(date[1])) {
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
    let c;
    // A script in b has a parent
    for (a = b.getElementsByTagName('script'); (c = a[0]);) c.parentNode!.removeChild(c);
    return b.innerHTML;
}

/**
 * Secure formula
 */
function secureFormula(oldValue: string) {
    let newValue = '';
    let inside = 0;

    // for...of would read the string by code points, and this loop reads its UTF-16 code units
    // eslint-disable-next-line @typescript-eslint/prefer-for-of
    for (let i = 0; i < oldValue.length; i++) {
        if (oldValue[i] === '"') {
            if (inside === 0) {
                inside = 1;
            } else {
                inside = 0;
            }
        }

        if (inside === 1) {
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
     * Get the td of a cell.
     * @param x - The column, or the cell name, for example "B3".
     * @param y - The row.
     */
    // A cell name, or the coordinates of the cell
    get(x: number | string, y?: number) {
        const worksheet = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x);

            x = cell[0];
            y = cell[1];
        }

        // y is set with the coordinates, and from the name
        return worksheet.records[y!][x].element;
    }

    /**
     * Get the td of a cell from its coordinates.
     * @param x - The column.
     * @param y - The row.
     */
    getFromCoords(x: number, y: number) {
        const worksheet = this.worksheet;

        return worksheet.records[y][x].element;
    }

    /**
     * Get the text that a cell shows.
     * @param x - The column, or the cell name, for example "B3".
     * @param y - The row.
     */
    // A cell name, or the coordinates of the cell
    getLabel(x: number | string, y?: number) {
        const worksheet = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x);

            x = cell[0];
            y = cell[1];
        }

        // y is set with the coordinates, and from the name
        return worksheet.records[y!][x].element.innerHTML;
    }

    /**
     * Show the value of a cell again, for example the new result of its formula, as update shows it. It writes no data
     * and dispatches no event.
     */
    showValue(x: number, y: number) {
        const worksheet = this.worksheet;
        let value = worksheet.options.data[y][x];
        // The cell shows the secured formula, as create and update show it
        if (`${value}`.substr(0, 1) === '=' && worksheet.options.secureFormulas === true) {
            value = secureFormula(value as string);
        }
        const c = this.getColumnIndex(x, y);
        this.display(x, y, value, false, c, worksheet.options.columns?.[c]);
    }

    /**
     * Show a value in a cell, as its column type shows it, and call the render of the column. c is the column whose
     * options the cell uses, and column its options. With write, the data of the cell gets the value too. It returns the value that the cell shows, which an image column changes to text.
     */
    private display(
        x: number,
        y: number,
        value: jspreadsheet.CellValue | undefined,
        write: boolean,
        c: number,
        column: Worksheet['options']['columns'][number] | undefined
    ) {
        const worksheet = this.worksheet;

        // A custom editor can have setValue, which is not in the public typings
        const editor = (column && typeof column.type === 'object' ? column.type : null) as
            (jspreadsheet.CustomEditor & { setValue?: (cell: HTMLTableCellElement, value: jspreadsheet.CellValue | undefined) => void }) | null;
        if (editor) {
            // Update data and cell
            if (write) {
                worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
            }
            if (typeof editor.setValue === 'function') {
                editor.setValue(worksheet.records[y][x].element, value);
            }
        } else {
            // Native functions
            if (column && (column.type === 'checkbox' || column.type === 'radio')) {
                // Unchecked all options
                if (column.type === 'radio' && write) {
                    for (const dataRow of worksheet.options.data) {
                        // The row source mode has rows without values
                        if (dataRow) {
                            dataRow[c] = false;
                        }
                    }
                }

                // Update data and cell
                (worksheet.records[y][x].element.children[0] as HTMLInputElement).checked =
                    // eslint-disable-next-line eqeqeq -- a checkbox is checked by 1, true and the texts "1", "true" and "TRUE"
                    value == 1 || value == true || value == 'true' || value == 'TRUE' ? true : false;
                if (write) {
                    worksheet.options.data[y][x] = (worksheet.records[y][x].element.children[0] as HTMLInputElement).checked;
                }
            } else if (column?.type === 'dropdown') {
                // Update data and cell
                if (write) {
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                }
                worksheet.records[y][x].element.textContent = this.getDropDownValue(c, value);
            } else if (column?.type === 'calendar') {
                // Update data and cell
                if (write) {
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                }
                worksheet.records[y][x].element.textContent = this.formatCalendarValue(value, column.options?.format);
            } else if (column?.type === 'color') {
                // Update color
                if (write) {
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                }
                // Render
                if (column.render === 'square') {
                    const color = document.createElement('div');
                    color.className = 'color';
                    color.style.backgroundColor = value as string;
                    worksheet.records[y][x].element.textContent = '';
                    worksheet.records[y][x].element.appendChild(color);
                } else {
                    worksheet.records[y][x].element.style.color = value as string;
                    worksheet.records[y][x].element.textContent = value as string;
                }
            } else if (column?.type === 'image') {
                value = `${value}`;
                if (write) {
                    worksheet.options.data[y][x] = value;
                }
                worksheet.records[y][x].element.innerHTML = '';
                if (value?.substr(0, 10) === 'data:image') {
                    const img = document.createElement('img');
                    img.src = value;
                    worksheet.records[y][x].element.appendChild(img);
                }
            } else {
                // Update data and cell
                if (write) {
                    worksheet.options.data[y][x] = value as jspreadsheet.CellValue;
                }
                // Label
                if (column?.type === 'html') {
                    worksheet.records[y][x].element.innerHTML = stripScript(this.parseValue(x, y, value, worksheet.records[y][x].element));
                } else {
                    if (worksheet.parent.config.parseHTML === true) {
                        worksheet.records[y][x].element.innerHTML = stripScript(this.parseValue(x, y, value, worksheet.records[y][x].element));
                    } else {
                        worksheet.records[y][x].element.textContent = this.parseValue(x, y, value, worksheet.records[y][x].element) as string;
                    }
                }
                // Handle big text inside a cell
                if (
                    column?.wordWrap !== false &&
                    (worksheet.options.wordWrap === true || column?.wordWrap === true || worksheet.records[y][x].element.innerHTML.length > 200)
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
                (worksheet.records[y]?.[x] ? worksheet.records[y][x].element : null) as HTMLTableCellElement,
                value,
                parseInt(x),
                parseInt(y),
                worksheet,
                column
            );
        }

        // Accessible name of the active cell
        if (worksheet.selection.activeCell && worksheet.records[y]?.[x]?.element === worksheet.selection.activeCell) {
            worksheet.accessibility.setCellLabel(worksheet.selection.activeCell);
        }

        return value;
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
        if (c === previous) {
            return;
        }

        const td = this.create(x, y, worksheet.options.data[y][x], true);

        function getOptions(i: number) {
            return worksheet.options.columns?.[i] || {};
        }
        const before = getOptions(previous);
        const after = getOptions(c);

        if ((before.readOnly === true) !== (after.readOnly === true)) {
            element.classList.toggle('readonly', after.readOnly === true);
        }
        if (before.align !== after.align) {
            element.style.textAlign = td.style.textAlign;
        }
        // create wraps the text with the wordWrap options, or when it is longer than 200 characters
        function wraps(options: jspreadsheet.Column, cell: HTMLElement) {
            return options.wordWrap !== false && (worksheet.options.wordWrap === true || options.wordWrap === true || cell.innerHTML.length > 200);
        }
        if (wraps(before, element) !== wraps(after, td)) {
            element.style.whiteSpace = td.style.whiteSpace;
        }
        if (before.type === 'color' || after.type === 'color') {
            element.style.color = td.style.color;
        }
        element.classList.toggle('red', td.classList.contains('red'));
        element.classList.toggle('jss_dropdown', td.classList.contains('jss_dropdown'));

        element.innerHTML = '';
        while (td.firstChild) {
            element.appendChild(td.firstChild);
        }

        if (c === x) {
            cellColumns.delete(element);
        } else {
            cellColumns.set(element, c - x);
        }

        // The cell has new content, which the handler can change as for a new cell
        worksheet.dispatch('oncreatecell', worksheet, element, x, y, worksheet.options.data[y][x]);
    }

    /**
     * Get dropdown value from key
     */
    private getDropDownValue(column: number, key: unknown) {
        const worksheet = this.worksheet;

        const value = [];

        if (worksheet.options.columns?.[column]?.source) {
            // Create array from source
            const combo = [];
            const source = worksheet.options.columns[column].source;

            for (const option of source) {
                if (typeof option === 'object') {
                    combo[option.id] = option.name;
                } else {
                    combo[option] = option;
                }
            }

            // Guarantee single multiple compatibility
            const delimiter = worksheet.options.columns[column].delimiter || ';';
            const keys = Array.isArray(key) ? key : `${key}`.split(delimiter);

            for (const optionKey of keys) {
                if (typeof optionKey === 'object') {
                    value.push(combo[optionKey.id]);
                } else {
                    if (combo[optionKey]) {
                        value.push(combo[optionKey]);
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
        const column = worksheet.options.columns?.[c];

        // Create cell and properties
        const td = document.createElement('td');
        if (c !== i) {
            cellColumns.set(td, c - parseInt(i));
        }
        td.setAttribute('data-x', i);
        td.setAttribute('data-y', j);
        td.setAttribute('aria-colindex', i + 2);

        if (worksheet.headers[i].style.display === 'none') {
            td.style.display = 'none';
        }
        // Security
        if (`${value}`.substr(0, 1) === '=' && worksheet.options.secureFormulas === true) {
            const val = secureFormula(value as string);
            if (val !== value) {
                // Update the data container
                value = val;
            }
        }

        // Text align, before the content, so that the formatValue hook can change it
        const colAlign = column?.align || worksheet.options.defaultColAlign || 'center';
        td.style.textAlign = colAlign;

        // Custom column
        if (column && typeof column.type === 'object') {
            if (worksheet.parent.config.parseHTML === true) {
                td.innerHTML = value as string;
            } else {
                td.textContent = value as string;
            }
            if (typeof column.type.createCell === 'function') {
                column.type.createCell(td, value, parseInt(i), parseInt(j), worksheet, column);
            }
        } else {
            // Hidden column
            if (column?.type === 'hidden') {
                td.style.display = 'none';
                td.textContent = value as string;
            } else if (column && (column.type === 'checkbox' || column.type === 'radio')) {
                // Create input
                const element = document.createElement('input');
                element.type = column.type;
                element.name = `c${c}`;
                // eslint-disable-next-line eqeqeq -- a checkbox is checked by 1, true and the texts "1" and "true"
                element.checked = value == 1 || value == true || value == 'true' ? true : false;
                element.onclick = () => {
                    worksheet.setValue(td, element.checked);
                };

                if (column.readOnly === true || worksheet.options.editable === false) {
                    element.setAttribute('disabled', 'disabled');
                }

                // Append to the table
                td.appendChild(element);
                // Make sure the values are correct
                worksheet.options.data[j][i] = element.checked;
            } else if (column?.type === 'calendar') {
                td.textContent = this.formatCalendarValue(value, column.options?.format);
            } else if (column?.type === 'dropdown') {
                // Create dropdown cell
                td.classList.add('jss_dropdown');
                td.textContent = this.getDropDownValue(c, value);
            } else if (column?.type === 'color') {
                if (column.render === 'square') {
                    const color = document.createElement('div');
                    color.className = 'color';
                    color.style.backgroundColor = value as string;
                    td.appendChild(color);
                } else {
                    td.style.color = value as string;
                    td.textContent = value as string;
                }
            } else if (column?.type === 'image') {
                // The value can be a number or a boolean
                if (value && `${value}`.substr(0, 10) === 'data:image') {
                    const img = document.createElement('img');
                    img.src = value as string;
                    td.appendChild(img);
                }
            } else {
                if (column?.type === 'html') {
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
        if (column?.readOnly === true) {
            td.className = 'readonly';
        }

        // Wrap option
        if (column?.wordWrap !== false && (worksheet.options.wordWrap === true || column?.wordWrap === true || td.innerHTML.length > 200)) {
            td.style.whiteSpace = 'pre-wrap';
        }

        // Overflow
        if (i > 0) {
            if (worksheet.options.textOverflow === true) {
                if (value || td.innerHTML) {
                    worksheet.records[j][i - 1].element.style.overflow = 'hidden';
                } else {
                    if (i === worksheet.options.columns.length - 1) {
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
     * Update the content of a cell.
     * @param x - The column.
     * @param y - The row.
     * @param value - The new value.
     * @param force - Also update a read-only cell.
     */
    // setValue without a value clears the cell, and the data keeps undefined. The DOM converts a number, a boolean or
    // undefined to text, so value is cast to text where the DOM takes text.
    update(x: number, y: number, value: jspreadsheet.CellValue | undefined, force?: boolean) {
        const worksheet = this.worksheet;

        // In the row source mode, a row outside the window has no record: the application keeps its values
        if (worksheet.rowSource && !worksheet.records[y]?.[x]) {
            return null;
        }

        // Column options
        const c = this.getColumnIndex(x, y);
        const column = worksheet.options.columns?.[c];

        let record: HistoryCellChange | null = null;

        // A read-only cell does not change, and has no record
        if (worksheet.records[y][x].element.classList.contains('readonly') === true && !force) {
            // Do nothing
        } else {
            // Security
            if (`${value}`.substr(0, 1) === '=' && worksheet.options.secureFormulas === true) {
                const val = secureFormula(value as string);
                if (val !== value) {
                    // Update the data container
                    value = val;
                }
            }

            // On change
            const val = worksheet.dispatch('onbeforechange', worksheet, worksheet.records[y][x].element, x, y, value);

            // False cancels the change, and another value is the new value
            if (val === false) {
                return null;
            }
            if (val != null) {
                value = val;
            }

            if (column && typeof column.type === 'object' && typeof column.type.updateCell === 'function') {
                const result = column.type.updateCell(worksheet.records[y][x].element, value, parseInt(x), parseInt(y), worksheet, column);

                if (result !== undefined) {
                    value = result;
                }
            }

            // History format
            const oldValue = worksheet.options.data[y][x];
            record = {
                x,
                y,
                col: x,
                row: y,
                value,
                oldValue,
            };

            value = this.display(x, y, value, true, c, column);

            // On change. updateCell has read worksheet.records[y][x] already, so the cell is there.
            worksheet.dispatch(
                'onchange',
                worksheet,
                (worksheet.records[y]?.[x] ? worksheet.records[y][x].element : null) as HTMLTableCellElement,
                x,
                y,
                value,
                oldValue
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

        if (worksheet.options.columns?.[x]?.type === 'hidden' && worksheet.options.mergeCells) {
            const merged = worksheet.options.mergeCells[getColumnNameFromId([x, y])];
            if (merged) {
                for (let i = x + 1; i < x + merged[0]; i++) {
                    if (worksheet.headers[i] && worksheet.headers[i].style.display !== 'none' && worksheet.options.columns[i]?.type !== 'hidden') {
                        if (
                            worksheet.options.columns[i] &&
                            (worksheet.options.columns[i].type === 'checkbox' || worksheet.options.columns[i].type === 'radio')
                        ) {
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
     * The text of a calendar cell: its date in the format of the column. A value that is not a date in the ISO format,
     * such as "05/03/2024", is read with the format of the column.
     * @param value - The value of the cell.
     * @param format - The format of the column, such as "DD/MM/YYYY".
     */
    formatCalendarValue(value: jspreadsheet.CellValue | undefined, format: string | undefined) {
        const date = validDate(value) ? value : parseDate(`${value ?? ''}`, format || 'YYYY-MM-DD');
        // A value that is not a date shows as it is
        return formatDate(date || value, format) || `${value ?? ''}`;
    }

    /**
     * The value that a cell shows: the result of a formula, then the text of the formatValue hook, or the value with the mask
     * of the column. The DOM converts a number, a boolean or undefined to text.
     */
    parseValue(i: number, j: number, value: jspreadsheet.CellValue | undefined, cell?: HTMLElement) {
        const worksheet = this.worksheet;

        if (`${value}`.substr(0, 1) === '=' && worksheet.parent.config.parseFormulas !== false) {
            // The cell shows the result as text, as the DOM converts it. An array from a range shows its items with commas.
            value = worksheet.formulas.execute(value as string, i, j) as jspreadsheet.CellValue;
        }

        // Column options
        const options = worksheet.options.columns?.[this.getColumnIndex(i, j)];

        // A hidden column shows its value when the cell is created, and the update keeps that
        const formatValue = worksheet.parent.config.formatValue;
        if (typeof formatValue === 'function' && options?.type !== 'hidden') {
            // A caller can give the coordinates as text
            const text = formatValue(worksheet, Number(i), Number(j), value, cell);
            if (text !== undefined) {
                // The color of a section of the mask does not apply to the text of the hook
                cell?.classList.remove('red');
                return text;
            }
        }
        if (options && !isFormula(value)) {
            // Mask options
            const opt = getMask(options);
            if (opt) {
                const masked = render(value, opt);
                value = masked.text;
                // The color of the section of the format, such as [Red]
                if (cell) {
                    if (masked.color === 'red') {
                        cell.classList.add('red');
                    } else {
                        cell.classList.remove('red');
                    }
                }
            }
        }

        return value;
    }
}
