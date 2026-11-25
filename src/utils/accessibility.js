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
