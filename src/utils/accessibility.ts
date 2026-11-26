import jSuites from 'jsuites';

import { getColumnNameFromId } from './internalHelpers.js';
import type { Spreadsheet, Worksheet } from '../types/internal';

/**
 * ARIA grid semantics of the worksheet table (https://www.w3.org/WAI/ARIA/apg/patterns/grid/).
 *
 * Lazy loading and pagination attach only some of the rows to the table, so the table gives its full size in
 * aria-rowcount and aria-colcount, and each row and cell gives its position in aria-rowindex and aria-colindex.
 * The rows of the thead come first, then the data rows, then the rows of the tfoot. Column 1 is the row number column,
 * so the cell (x, y) has the aria-colindex x + 2. A hidden column is not in the DOM order, so each cell has an
 * explicit aria-colindex. A data cell gets the gridcell role from the grid and its row.
 */

/**
 * English texts of the accessible names. The spreadsheet option labels gives other texts, for example in the language
 * of the page. Without the option, jspreadsheet.setDictionary can translate the English texts. {0} in a text stands for
 * a value, such as the address of a cell.
 */
const defaultLabels = {
    selectAll: 'Select all',
    editCell: 'Edit cell {0}',
    clipboardHelper: 'Clipboard helper',
};

/**
 * Text of an accessible name
 */
export const getLabelText = function (this: Spreadsheet, key: keyof typeof defaultLabels, value?: string) {
    const spreadsheet = this;

    const labels = spreadsheet && spreadsheet.config && spreadsheet.config.labels;
    const text = labels && typeof labels[key] === 'string' ? labels[key] : jSuites.translate(defaultLabels[key]);

    return value === undefined ? text : text.split('{0}').join(value);
};

/**
 * Number of rows before the first data row
 */
const getHeaderRows = function (this: Worksheet) {
    const obj = this;

    return obj.thead.children.length;
};

/**
 * Roles of the table, its sections and its header rows
 */
export const setGridRoles = function (this: Worksheet) {
    const obj = this;

    obj.table.setAttribute('role', 'grid');
    // A selection can have more than one cell
    obj.table.setAttribute('aria-multiselectable', 'true');
    obj.thead.setAttribute('role', 'rowgroup');
    obj.tbody.setAttribute('role', 'rowgroup');

    for (let j = 0; j < obj.thead.children.length; j++) {
        const tr = obj.thead.children[j];
        tr.setAttribute('role', 'row');
        tr.setAttribute('aria-rowindex', j + 1);

        // Nested headers, after the select all corner
        if (tr.classList.contains('jss_nested')) {
            for (let i = 1; i < tr.children.length; i++) {
                tr.children[i].setAttribute('role', 'columnheader');
            }
        }
    }

    // Select all corner
    obj.headerContainer.children[0].setAttribute('aria-colindex', 1);
};

/**
 * Accessible names of the table, the select all corner and the clipboard helper
 */
export const setGridLabels = function (this: Worksheet) {
    const obj = this;

    if (obj.options.worksheetName) {
        obj.table.setAttribute('aria-label', obj.options.worksheetName);
    }

    obj.headerContainer.children[0].setAttribute('aria-label', getLabelText.call(obj.parent, 'selectAll'));

    obj.clipboard.textarea.setAttribute('aria-label', getLabelText.call(obj.parent, 'clipboardHelper'));
};

/**
 * Role and position of a column header
 */
export const setHeaderRole = function (td: HTMLTableCellElement, x: number) {
    td.setAttribute('role', 'columnheader');
    td.setAttribute('aria-colindex', parseInt(x) + 2);
};

/**
 * Role and position of a data row and its row number
 */
export const setRowRole = function (this: Worksheet, tr: HTMLTableRowElement, y: number) {
    const obj = this;

    tr.setAttribute('role', 'row');
    updateRowIndex.call(obj, tr, y);
    tr.children[0].setAttribute('role', 'rowheader');
    tr.children[0].setAttribute('aria-colindex', 1);
};

/**
 * Position of a data row after a change of its index
 */
export const updateRowIndex = function (this: Worksheet, tr: HTMLTableRowElement, y: number) {
    const obj = this;

    tr.setAttribute('aria-rowindex', getHeaderRows.call(obj) + parseInt(y) + 1);
};

/**
 * Column positions of the cells of a filter or footer row. The first cell is in the row number column.
 */
const updateRowCells = function (tr: Element) {
    for (let i = 0; i < tr.children.length; i++) {
        tr.children[i].setAttribute('aria-colindex', i + 1);
    }
};

/**
 * Size of the grid, and the positions in the nested header, filter and footer rows, after a change of the rows or
 * columns. These rows have few cells.
 */
export const updateGridSize = function (this: Worksheet) {
    const obj = this;

    const headerRows = getHeaderRows.call(obj);
    let footerRows = 0;

    for (let j = 0; j < headerRows; j++) {
        const tr = obj.thead.children[j];
        if (tr.classList.contains('jss_nested')) {
            // A nested header starts at its first column, which is in data-column
            tr.children[0].setAttribute('aria-colindex', 1);
            for (let i = 1; i < tr.children.length; i++) {
                tr.children[i].setAttribute('aria-colindex', parseInt(tr.children[i].getAttribute('data-column')) + 2);
            }
        } else if (tr !== obj.headerContainer) {
            updateRowCells(tr);
        }
    }

    if (obj.tfoot) {
        obj.tfoot.setAttribute('role', 'rowgroup');
        footerRows = obj.tfoot.children.length;
        for (let j = 0; j < footerRows; j++) {
            obj.tfoot.children[j].setAttribute('role', 'row');
            obj.tfoot.children[j].setAttribute('aria-rowindex', headerRows + obj.rows.length + j + 1);
            updateRowCells(obj.tfoot.children[j]);
        }
    }

    obj.table.setAttribute('aria-rowcount', headerRows + obj.rows.length + footerRows);
    obj.table.setAttribute('aria-colcount', obj.headers.length + 1);
};

/**
 * Accessible name of the active cell: its address and its text. A cell without text, for example a checkbox or a color
 * square, gives its value. An image gives no value, because its value is the image data. A cell with an open editor has
 * no name, because the editor shows other text and has its own name.
 */
export const setCellLabel = function (this: Worksheet, td: HTMLTableCellElement) {
    const obj = this;

    if (td.classList.contains('editor') || (obj.editor.isOpen() && obj.editor.cell === td)) {
        td.removeAttribute('aria-label');
        return;
    }

    const x = parseInt(td.getAttribute('data-x'));
    const y = parseInt(td.getAttribute('data-y'));
    let text = td.textContent;
    if (!text && !td.querySelector('img')) {
        const value = obj.getValueFromCoords(x, y);
        text = value === null || value === undefined ? '' : '' + value;
    }

    td.setAttribute('aria-label', getColumnNameFromId([x, y]) + (text ? ' ' + text : ''));
};
