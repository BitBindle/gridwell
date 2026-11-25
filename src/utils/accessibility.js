import jSuites from 'jsuites';

import libraryBase from './libraryBase.js';
import { getColumnNameFromId } from './internalHelpers.js';

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
export const getLabelText = function (key, value) {
    const spreadsheet = this;

    const labels = spreadsheet && spreadsheet.config && spreadsheet.config.labels;
    const text = labels && typeof labels[key] === 'string' ? labels[key] : jSuites.translate(defaultLabels[key]);

    return value === undefined ? text : text.split('{0}').join(value);
};

/**
 * Number of rows before the first data row
 */
const getHeaderRows = function () {
    const obj = this;

    return obj.thead.children.length;
};

/**
 * Roles of the table, its sections and its header rows
 */
export const setGridRoles = function () {
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
export const setGridLabels = function () {
    const obj = this;

    if (obj.options.worksheetName) {
        obj.table.setAttribute('aria-label', obj.options.worksheetName);
    }

    obj.headerContainer.children[0].setAttribute('aria-label', getLabelText.call(obj.parent, 'selectAll'));

    obj.textarea.setAttribute('aria-label', getLabelText.call(obj.parent, 'clipboardHelper'));
};

/**
 * Role and position of a column header
 */
export const setHeaderRole = function (td, x) {
    td.setAttribute('role', 'columnheader');
    td.setAttribute('aria-colindex', parseInt(x) + 2);
};

/**
 * Role and position of a data row and its row number
 */
export const setRowRole = function (tr, y) {
    const obj = this;

    tr.setAttribute('role', 'row');
    updateRowIndex.call(obj, tr, y);
    tr.children[0].setAttribute('role', 'rowheader');
    tr.children[0].setAttribute('aria-colindex', 1);
};

/**
 * Position of a data row after a change of its index
 */
export const updateRowIndex = function (tr, y) {
    const obj = this;

    tr.setAttribute('aria-rowindex', getHeaderRows.call(obj) + parseInt(y) + 1);
};

/**
 * Column positions of the cells of a filter or footer row. The first cell is in the row number column.
 */
const updateRowCells = function (tr) {
    for (let i = 0; i < tr.children.length; i++) {
        tr.children[i].setAttribute('aria-colindex', i + 1);
    }
};

/**
 * Size of the grid, and the positions in the nested header, filter and footer rows, after a change of the rows or
 * columns. These rows have few cells.
 */
export const updateGridSize = function () {
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
export const setCellLabel = function (td) {
    const obj = this;

    if (td.classList.contains('editor') || (obj.edition && obj.edition[0] === td)) {
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

/**
 * A cell can be the tab stop when its row is in the table and the cell and its row are not hidden
 */
const isVisibleCell = function (td) {
    const obj = this;

    return !!td && !!td.parentNode && td.parentNode.parentNode === obj.tbody && td.style.display !== 'none' && td.parentNode.style.display !== 'none';
};

/**
 * First visible cell of the rows in the table
 */
const getFirstVisibleCell = function () {
    const obj = this;

    for (let j = 0; j < obj.tbody.children.length; j++) {
        const tr = obj.tbody.children[j];
        if (tr.style.display !== 'none') {
            // The first child is the row number
            for (let i = 1; i < tr.children.length; i++) {
                if (tr.children[i].style.display !== 'none') {
                    return tr.children[i];
                }
            }
        }
    }

    return null;
};

/**
 * Roving tabindex: the active cell is the only cell with tabindex 0. The grid is multiselectable, so the active cell
 * gives aria-selected="false" when it is not selected.
 */
const setActiveCell = function (td) {
    const obj = this;

    const previous = obj.activeCell;
    if (previous && previous !== td) {
        previous.tabIndex = -1;
        previous.removeAttribute('aria-label');
        if (previous.getAttribute('aria-selected') === 'false') {
            previous.removeAttribute('aria-selected');
        }
    }

    obj.activeCell = td;

    if (td) {
        // The editor of a cell takes the focus, so the cell is not a focus target while the editor is open
        if (!obj.edition || obj.edition[0] !== td) {
            td.tabIndex = 0;
        }
        setCellLabel.call(obj, td);
        if (td.getAttribute('aria-selected') !== 'true') {
            td.setAttribute('aria-selected', 'false');
        }
    }
};

/**
 * The focused element in the document or in the shadow root of the worksheet
 */
const getActiveElement = function () {
    const obj = this;

    const root = obj.table.getRootNode();

    return root.activeElement || null;
};

/**
 * The focus is on no element, for example after the focused element was removed
 */
const isFocusLost = function () {
    const active = document.activeElement;

    return !active || active === document.body;
};

/**
 * The tab stop of the grid is the selected cell. When no cell is selected, or the row of the selected cell is not in
 * the table, for example after lazy loading, pagination, a search or deleteRow, the active cell stays the tab stop if it
 * is still in the table, else the first visible cell is the tab stop. When the focused cell left the table, or the table
 * attached its row again, the focus fell to the body, and the tab stop gets it back. A focused cell that is no longer the
 * tab stop also gives the focus to the tab stop.
 */
export const updateTabStop = function () {
    const obj = this;

    let td = null;
    if (obj.selectedCell) {
        const x = obj.selectedCell[0];
        const y = obj.selectedCell[1];
        td = obj.records[y] && obj.records[y][x] ? obj.records[y][x].element : null;
    }

    if (!isVisibleCell.call(obj, td)) {
        td = isVisibleCell.call(obj, obj.activeCell) ? obj.activeCell : getFirstVisibleCell.call(obj);
    }

    setActiveCell.call(obj, td);

    if (td && !obj.edition && libraryBase.jspreadsheet.current === obj) {
        const active = getActiveElement.call(obj);
        // The focus was in the grid at the last selection
        const lost = obj.focusInGrid === true && isFocusLost();
        const otherCell = !!active && active !== td && active.tagName === 'TD' && obj.tbody.contains(active);
        if (lost || otherCell) {
            td.focus({ preventScroll: true });
        }
    }
};

/**
 * When the focus goes from the grid to an element outside the spreadsheet, a later change of the grid does not take it back
 */
export const setGridFocusEvents = function () {
    const obj = this;

    obj.table.addEventListener('focusout', function (e) {
        if (e.relatedTarget && e.relatedTarget.nodeType === 1 && !obj.parent.element.contains(e.relatedTarget)) {
            obj.focusInGrid = false;
        }
    });
};

/**
 * The selected cell is the tab stop of the grid. It also gets the focus when the focus is in the grid, after a press on
 * the grid, or when the focus fell from the grid to the body, for example after lazy loading removed the focused row.
 * A selection does not take the focus from other parts of the page.
 */
export const updateActiveCell = function (x, y, origin) {
    const obj = this;

    const td = obj.records[y] && obj.records[y][x] && obj.records[y][x].element;
    if (!td) {
        return;
    }

    setActiveCell.call(obj, td);

    const active = getActiveElement.call(obj);
    const pointer = origin && origin.target && origin.target.nodeType === 1 && obj.table.contains(origin.target);
    const inGrid = pointer || (!!active && (active === obj.parent.element || active === obj.textarea || obj.table.contains(active)));
    const dropped = isFocusLost() && libraryBase.jspreadsheet.current === obj;
    // The focus was in the grid at the last selection
    obj.focusInGrid = inGrid || (dropped && obj.focusInGrid === true);

    if (obj.focusInGrid && !obj.edition && active !== td) {
        td.focus({ preventScroll: true });
    }
};

/**
 * After the editor closes, the cell is a focus target again. It gets the focus back when the focus fell to the body
 * with the removed editor.
 */
export const restoreActiveCell = function (cell) {
    const obj = this;

    if (obj.activeCell !== cell) {
        cell.tabIndex = -1;
        return;
    }

    cell.tabIndex = 0;
    setCellLabel.call(obj, cell);

    setTimeout(function () {
        if (isFocusLost() && libraryBase.jspreadsheet.current === obj && obj.activeCell === cell && !obj.edition) {
            cell.focus({ preventScroll: true });
        }
    });
};
