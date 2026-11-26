import jSuites from 'jsuites';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import type { Worksheet } from '../types/internal';

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
 * Role and position of a column header
 */
export function setHeaderRole(td: HTMLTableCellElement, x: number) {
    td.setAttribute('role', 'columnheader');
    td.setAttribute('aria-colindex', parseInt(x) + 2);
}

/**
 * Column positions of the cells of a filter or footer row. The first cell is in the row number column.
 */
function updateRowCells(tr: Element) {
    for (let i = 0; i < tr.children.length; i++) {
        tr.children[i].setAttribute('aria-colindex', i + 1);
    }
}

/**
 * ARIA grid semantics of the worksheet table (https://www.w3.org/WAI/ARIA/apg/patterns/grid/): the roles, the accessible
 * names and the size of the grid.
 *
 * Lazy loading and pagination attach only some of the rows to the table, so the table gives its full size in
 * aria-rowcount and aria-colcount, and each row and cell gives its position in aria-rowindex and aria-colindex.
 * The rows of the thead come first, then the data rows, then the rows of the tfoot. Column 1 is the row number column,
 * so the cell (x, y) has the aria-colindex x + 2. A hidden column is not in the DOM order, so each cell has an
 * explicit aria-colindex. A data cell gets the gridcell role from the grid and its row.
 *
 * Each worksheet has one, as worksheet.accessibility.
 */
export class Accessibility {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Text of an accessible name
     */
    getLabelText(key: keyof typeof defaultLabels, value?: string) {
        const spreadsheet = this.worksheet.parent;

        const labels = spreadsheet?.config?.labels;
        const text = labels && typeof labels[key] === 'string' ? labels[key] : jSuites.translate(defaultLabels[key]);

        return value === undefined ? text : text.split('{0}').join(value);
    }

    /**
     * Number of rows before the first data row
     */
    private getHeaderRows() {
        const worksheet = this.worksheet;

        return worksheet.thead.children.length;
    }

    /**
     * Roles of the table, its sections and its header rows
     */
    setGridRoles() {
        const worksheet = this.worksheet;

        worksheet.table.setAttribute('role', 'grid');
        // A selection can have more than one cell
        worksheet.table.setAttribute('aria-multiselectable', 'true');
        worksheet.thead.setAttribute('role', 'rowgroup');
        worksheet.tbody.setAttribute('role', 'rowgroup');

        for (let j = 0; j < worksheet.thead.children.length; j++) {
            const tr = worksheet.thead.children[j];
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
        worksheet.headerContainer.children[0].setAttribute('aria-colindex', 1);
    }

    /**
     * Accessible names of the table, the select all corner and the clipboard helper
     */
    setGridLabels() {
        const worksheet = this.worksheet;

        if (worksheet.options.worksheetName) {
            worksheet.table.setAttribute('aria-label', worksheet.options.worksheetName);
        }

        worksheet.headerContainer.children[0].setAttribute('aria-label', this.getLabelText('selectAll'));

        worksheet.clipboard.textarea.setAttribute('aria-label', this.getLabelText('clipboardHelper'));
    }

    /**
     * Role and position of a data row and its row number
     */
    setRowRole(tr: HTMLTableRowElement, y: number) {
        const worksheet = this.worksheet;

        tr.setAttribute('role', 'row');
        this.updateRowIndex(tr, y);
        tr.children[0].setAttribute('role', 'rowheader');
        tr.children[0].setAttribute('aria-colindex', 1);
    }

    /**
     * Position of a data row after a change of its index
     */
    updateRowIndex(tr: HTMLTableRowElement, y: number) {
        const worksheet = this.worksheet;

        tr.setAttribute('aria-rowindex', this.getHeaderRows() + parseInt(y) + 1);
    }

    /**
     * Size of the grid, and the positions in the nested header, filter and footer rows, after a change of the rows or
     * columns. These rows have few cells.
     */
    updateGridSize() {
        const worksheet = this.worksheet;

        const headerRows = this.getHeaderRows();
        let footerRows = 0;

        for (let j = 0; j < headerRows; j++) {
            const tr = worksheet.thead.children[j];
            if (tr.classList.contains('jss_nested')) {
                // A nested header starts at its first column, which is in data-column
                tr.children[0].setAttribute('aria-colindex', 1);
                for (let i = 1; i < tr.children.length; i++) {
                    tr.children[i].setAttribute('aria-colindex', parseInt(tr.children[i].getAttribute('data-column')) + 2);
                }
            } else if (tr !== worksheet.headerContainer) {
                updateRowCells(tr);
            }
        }

        if (worksheet.tfoot) {
            worksheet.tfoot.setAttribute('role', 'rowgroup');
            footerRows = worksheet.tfoot.children.length;
            for (let j = 0; j < footerRows; j++) {
                worksheet.tfoot.children[j].setAttribute('role', 'row');
                worksheet.tfoot.children[j].setAttribute('aria-rowindex', headerRows + worksheet.rows.length + j + 1);
                updateRowCells(worksheet.tfoot.children[j]);
            }
        }

        worksheet.table.setAttribute('aria-rowcount', headerRows + worksheet.rows.length + footerRows);
        worksheet.table.setAttribute('aria-colcount', worksheet.headers.length + 1);
    }

    /**
     * Accessible name of the active cell: its address and its text. A cell without text, for example a checkbox or a color
     * square, gives its value. An image gives no value, because its value is the image data. A cell with an open editor has
     * no name, because the editor shows other text and has its own name.
     */
    setCellLabel(td: HTMLTableCellElement) {
        const worksheet = this.worksheet;

        if (td.classList.contains('editor') || (worksheet.editor.isOpen() && worksheet.editor.cell === td)) {
            td.removeAttribute('aria-label');
            return;
        }

        const x = parseInt(td.getAttribute('data-x'));
        const y = parseInt(td.getAttribute('data-y'));
        let text = td.textContent;
        if (!text && !td.querySelector('img')) {
            const value = worksheet.getValueFromCoords(x, y);
            text = value === null || value === undefined ? '' : `${value}`;
        }

        td.setAttribute('aria-label', getColumnNameFromId([x, y]) + (text ? ` ${text}` : ''));
    }
}
