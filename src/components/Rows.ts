import { translate } from '../utils/dictionary.js';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, WorksheetHistoryRecord } from '../types/internal';
import { toGrid } from '../utils/zoom.js';
import { askConfirm } from '../utils/dialogs.js';

/**
 * The rows of a worksheet: it creates, inserts, moves, deletes, hides and shows them, gets and sets their heights and
 * data, and hides or shows the index column.
 *
 * Each worksheet has one, as worksheet.rowManager. worksheet.rows is the list of row records, which applications use.
 */
export class Rows {
    readonly worksheet: Worksheet;

    /** The elements of the frozen rows, in their order. */
    private frozen: HTMLTableRowElement[] = [];

    /** Places the frozen rows again when the header row or a frozen row changes its height. */
    private freezeObserver: ResizeObserver | null = null;

    /** True while a placement of the frozen rows waits for the end of the current task. */
    private placing = false;

    /** Places the header rows again when a header row changes its height. Only a header of more than one row has it. */
    private headerObserver: ResizeObserver | null = null;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Change the number of frozen rows and frozen columns, and apply them at once. Undefined keeps the current number. A
     * value that is not a finite number of 0 or more gives 0.
     */
    setFreeze(rows?: number, columns?: number) {
        const worksheet = this.worksheet;
        function count(value: number) {
            const number = Math.floor(Number(value));
            return Number.isFinite(number) && number > 0 ? number : 0;
        }
        if (rows !== undefined) {
            worksheet.options.freezeRows = count(rows);
        }
        if (columns !== undefined) {
            worksheet.options.freezeColumns = count(columns);
        }
        // The row source mode makes its rows again, with the new frozen rows first. A column window makes its columns
        // again, with the new frozen columns first.
        if (worksheet.rowSource && (rows !== undefined || (columns !== undefined && worksheet.rowSource.columns))) {
            worksheet.rowSource.reset();
        }
        this.updateFreeze();
        // The scroll handler keeps this value only while columns are frozen, so it can be old
        worksheet.lastScrollLeft = worksheet.content.scrollLeft;
        worksheet.columns.updateFreezePosition();
        // With lazy loading, the selected row can leave the table
        worksheet.selection.updateTabStop();
    }

    /**
     * The number of frozen rows: the freezeRows option, but not more than the rows of the worksheet. With lazy loading,
     * at most 40 rows freeze: a load starts 30 rows from the edge of a page of 100 rows and removes 31 rows, so the page
     * needs 60 rows that scroll.
     */
    getFreezeCount() {
        const worksheet = this.worksheet;
        const max = worksheet.options.lazyLoading === true ? Math.min(40, worksheet.rows.length) : worksheet.rows.length;
        return Math.max(0, Math.min(Math.floor(worksheet.options.freezeRows || 0), max));
    }

    /**
     * The height that the sticky header covers at the top of the content. Each header row sticks below the rows above
     * it, so the header covers the height of all its rows.
     */
    getHeaderHeight() {
        let height = 0;
        for (const row of Array.from(this.worksheet.thead.children) as HTMLElement[]) {
            height += row.offsetHeight;
        }
        return height;
    }

    /**
     * Give each header row its distance from the top of the content in the custom property --jss-header-top: the height
     * of the header rows above it. The nested headers and the filters stick below each other, and do not cover each
     * other. A header of one row needs no distance.
     */
    placeHeader() {
        const rows = Array.from(this.worksheet.thead.children) as HTMLElement[];
        if (rows.length < 2) {
            return;
        }

        // All reads come before the writes, so the browser computes the layout once
        const heights = rows.map((row) => row.offsetHeight);
        let top = 0;
        rows.forEach((row, j) => {
            const value = `${top}px`;
            if (row.style.getPropertyValue('--jss-header-top') !== value) {
                row.style.setProperty('--jss-header-top', value);
            }
            top += heights[j];
        });
    }

    /**
     * Place the header rows when the worksheet shows, and again when a header row changes its height, for example after
     * a web font loads. The frozen rows and the fill handle depend on the height of the header, so they follow.
     */
    observeHeader() {
        const rows = Array.from(this.worksheet.thead.children);
        if (rows.length < 2) {
            return;
        }
        if (typeof ResizeObserver !== 'function') {
            queueMicrotask(() => this.placeHeader());
            return;
        }
        this.headerObserver?.disconnect();
        this.headerObserver = new ResizeObserver(() => {
            this.placeHeader();
            if (this.frozen.length) {
                this.placeFrozen();
            }
            this.worksheet.selection.updateCornerCover();
        });
        for (const row of rows) {
            this.headerObserver.observe(row);
        }
    }

    /**
     * The height of the frozen rows that show.
     */
    getFreezeHeight() {
        let height = 0;
        for (const row of this.frozen) {
            if (row.parentNode === this.worksheet.tbody) {
                height += row.offsetHeight;
            }
        }
        return height;
    }

    /**
     * Make the first freezeRows rows stick below the header row. Each frozen row gets the class jss_frozen_row and its
     * distance from the top of the content in the custom property --jss-frozen-top, so its cells keep their own style.
     * Call it after the order or the number of the rows changes.
     */
    updateFreeze() {
        const worksheet = this.worksheet;

        const count = this.getFreezeCount();
        // A worksheet without frozen rows does no work, also no layout
        if (!count && !this.frozen.length) {
            return;
        }

        const frozen: HTMLTableRowElement[] = [];
        for (let j = 0; j < count; j++) {
            frozen.push(worksheet.rows[j].element);
        }
        const changed = frozen.length !== this.frozen.length || frozen.some((row, j) => row !== this.frozen[j]);

        // A row that is not in the first rows any more is not frozen
        const left = this.frozen.filter((row) => !frozen.includes(row));
        for (const row of left) {
            row.classList.remove('jss_frozen_row', 'jss_frozen_last');
            row.style.removeProperty('--jss-frozen-top');
        }
        for (const row of frozen) {
            row.classList.add('jss_frozen_row');
        }
        this.frozen = frozen;

        // A change can take a frozen row out of the rows that lazy loading attached, or put a row that left the frozen
        // rows ahead of the rows that scroll
        worksheet.view.attachFrozen(left);

        if (changed) {
            this.schedulePlace();
            this.observeFrozen();
        }
    }

    /**
     * Give each frozen row its distance from the top of the content. The last frozen row that shows gets the class
     * jss_frozen_last, which draws the line of the frozen pane.
     */
    placeFrozen() {
        const tbody = this.worksheet.tbody;

        // All reads come before the writes, so the browser computes the layout once
        const header = this.getHeaderHeight();
        const heights = this.frozen.map((row) => (row.style.display !== 'none' && row.parentNode === tbody ? row.offsetHeight : null));

        let top = header;
        let last: HTMLTableRowElement | null = null;
        this.frozen.forEach((row, j) => {
            const value = `${top}px`;
            if (row.style.getPropertyValue('--jss-frozen-top') !== value) {
                row.style.setProperty('--jss-frozen-top', value);
            }
            if (heights[j] !== null) {
                top += heights[j]!;
                last = row;
            }
        });
        for (const row of this.frozen) {
            if (row.classList.contains('jss_frozen_last') !== (row === last)) {
                row.classList.toggle('jss_frozen_last', row === last);
            }
        }
    }

    /**
     * Place the frozen rows at the end of the current task. The placement reads the layout, so many changes in one task,
     * for example the inserts of a batch, need one layout, not one for each change.
     */
    private schedulePlace() {
        if (this.placing) {
            return;
        }
        this.placing = true;
        queueMicrotask(() => {
            this.placing = false;
            // loadWorksheet can replace the rows of the worksheet before the placement
            if (this.worksheet.rowManager !== this) {
                return;
            }
            this.placeFrozen();

            // A cell of a frozen row and a frozen column sticks on both axes. The placement reads the scroll position.
            if (this.worksheet.options.freezeColumns) {
                this.worksheet.columns.updateFreezePosition();
            }
        });
    }

    /**
     * Observe the size of the header row and of the frozen rows. A hidden worksheet has no layout, and the observer places
     * the rows when the worksheet shows. The callback observes no new element, so it cannot run in a loop.
     */
    private observeFrozen() {
        if (typeof ResizeObserver !== 'function') {
            return;
        }
        if (!this.freezeObserver) {
            this.freezeObserver = new ResizeObserver(() => {
                this.placeFrozen();
            });
        }
        this.freezeObserver.disconnect();
        if (this.frozen.length) {
            this.freezeObserver.observe(this.worksheet.thead);
            for (const row of this.frozen) {
                this.freezeObserver.observe(row);
            }
        }
    }

    /**
     * Stop the size observers of the frozen rows and of the header rows, for example when the spreadsheet is destroyed.
     */
    stopFreeze() {
        this.freezeObserver?.disconnect();
        this.freezeObserver = null;
        this.headerObserver?.disconnect();
        this.headerObserver = null;
    }

    /**
     * Make the cell of column i of row j and its record, and call the render of the column. Returns the cell.
     */
    createRecord(j: number, i: number, value: jspreadsheet.CellValue) {
        const worksheet = this.worksheet;
        const element = worksheet.cells.create(i, j, value);
        worksheet.records[j][i] = { element, x: i, y: j };
        const column = worksheet.options.columns?.[i];
        if (column && typeof column.render === 'function') {
            column.render(element, value, parseInt(i), parseInt(j), worksheet, column);
        }
        return element;
    }

    /**
     * Create row
     */
    create(j: number, data: jspreadsheet.CellValue[]) {
        const worksheet = this.worksheet;

        // Create container
        if (!worksheet.records[j]) {
            worksheet.records[j] = [];
        }
        // Default data
        if (!data) {
            data = worksheet.options.data[j];
        }
        // New line of data to be append in the table
        const row = {
            element: document.createElement('tr'),
            y: j,
        };

        worksheet.rows[j] = row;

        row.element.setAttribute('data-y', j);
        // Index
        let index = null;

        // Set default row height
        if (worksheet.options.defaultRowHeight) {
            row.element.style.height = `${worksheet.options.defaultRowHeight}px`;
        }

        // Definitions
        if (worksheet.options.rows?.[j]) {
            if (worksheet.options.rows[j].height) {
                row.element.style.height = worksheet.options.rows[j].height as string;
            }
            if (worksheet.options.rows[j].title) {
                index = worksheet.options.rows[j].title;
            }
        }
        if (!index) {
            index = parseInt(j + 1);
        }
        // Row number label
        const td = document.createElement('td');
        // A title, or the number of the row, which the DOM writes as text
        td.innerHTML = index as string;
        td.setAttribute('data-y', j);
        td.className = 'jss_row';
        row.element.appendChild(td);

        // With a column window, the row has cells only for the frozen columns and the window, and two spacer cells
        const columns = worksheet.rowSource?.columns;
        if (columns) {
            columns.appendCells(row.element, (i) => row.element.appendChild(this.createRecord(j, i, data[i])));
        } else {
            const numberOfColumns = worksheet.columns.count();
            for (let i = 0; i < numberOfColumns; i++) {
                row.element.appendChild(this.createRecord(j, i, data[i]));
            }
        }

        // Role and position for assistive technology
        worksheet.accessibility.setRowRole(row.element, j);

        // Add row to the table body
        return row;
    }

    /**
     * The rows that insertRow inserts, with their position and their data, for the events.
     */
    private getInserted(numOfRows: number, rowNumber: number, insertBefore: boolean, data: jspreadsheet.CellValue[]) {
        const worksheet = this.worksheet;

        const onbeforeinsertrowRecords = [];

        for (let row = 0; row < numOfRows; row++) {
            const newRow = [];

            for (let col = 0; col < worksheet.options.columns.length; col++) {
                newRow[col] = data[col] ? data[col] : '';
            }

            onbeforeinsertrowRecords.push({
                row: row + rowNumber + (insertBefore ? 0 : 1),
                data: newRow,
            });
        }

        return onbeforeinsertrowRecords;
    }

    /**
     * Create the new rows with the data or empty, and put them in the table before the rows that follow them.
     * @param currentRows - The rows after the new rows, which insertRow took out of the list of rows.
     */
    private add(rowIndex: number, numOfRows: number, rowNumber: number, data: jspreadsheet.CellValue[], currentRows: Worksheet['rows']) {
        const worksheet = this.worksheet;

        const rowRecords = [];
        const rowData = [];
        const rowNode = [];

        for (let row = rowIndex; row < numOfRows + rowIndex; row++) {
            // Push data to the data container
            worksheet.options.data[row] = [];
            for (let col = 0; col < worksheet.options.columns.length; col++) {
                worksheet.options.data[row][col] = data[col] ? data[col] : '';
            }
            // Create row
            const newRow = this.create(row, worksheet.options.data[row]);
            // Append node
            if (currentRows[0]) {
                if (currentRows[0].element.parentNode === worksheet.tbody) {
                    worksheet.tbody.insertBefore(newRow.element, currentRows[0].element);
                }
            } else {
                if (worksheet.rows[rowNumber].element.parentNode === worksheet.tbody) {
                    worksheet.tbody.appendChild(newRow.element);
                }
            }
            // Record History
            rowRecords.push([...worksheet.records[row]]);
            rowData.push([...worksheet.options.data[row]]);
            rowNode.push(newRow);
        }

        return { rowRecords, rowData, rowNode };
    }

    /**
     * Insert rows.
     * @param mixed - The number of rows, or the data of one new row.
     * @param rowNumber - The row next to which the rows go. The last row by default.
     * @param insertBefore - Put the rows before that row instead of after it.
     */
    insert(mixed?: number | jspreadsheet.CellValue[], rowNumber?: number, insertBefore?: boolean | number): false | undefined {
        const worksheet = this.worksheet;

        // The application changes the rows of the row source mode, and calls resetRows
        if (worksheet.rowSource) {
            return false;
        }

        // Configuration
        if (worksheet.options.allowInsertRow !== false) {
            // Data to be insert
            let data: jspreadsheet.CellValue[] = [];

            // The insert could be lead by number of rows or the array of data
            let numOfRows;

            if (!Array.isArray(mixed)) {
                numOfRows = typeof mixed !== 'undefined' ? mixed : 1;
            } else {
                numOfRows = 1;

                if (mixed) {
                    data = mixed;
                }
            }

            // Direction
            insertBefore = insertBefore ? true : false;

            // Current column number
            const lastRow = worksheet.options.data.length - 1;

            if (rowNumber == null || rowNumber >= parseInt(lastRow) || rowNumber < 0) {
                rowNumber = lastRow;
            }

            const onbeforeinsertrowRecords = this.getInserted(numOfRows, rowNumber, insertBefore, data);

            // Onbeforeinsertrow
            if (worksheet.dispatch('onbeforeinsertrow', worksheet, onbeforeinsertrowRecords) === false) {
                return false;
            }

            // The merged ranges that the new rows split are removed. The other ranges stay and move with their cells.
            const crossed = worksheet.merges.isRowMerged(rowNumber, insertBefore);
            if (crossed.length && !askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
                return false;
            }

            // Clear any search
            if (worksheet.options.search === true) {
                if (worksheet.view.results && worksheet.view.results.length !== worksheet.rows.length) {
                    if (askConfirm(worksheet.parent.config, translate('This action will clear your search results. Are you sure?'))) {
                        worksheet.resetSearch();
                    } else {
                        return false;
                    }
                }

                worksheet.view.results = null;
            }

            // After the questions, so that a cancel keeps the ranges
            const merges = worksheet.merges.removeCrossed(crossed);

            // Insertbefore
            const rowIndex = !insertBefore ? rowNumber + 1 : rowNumber;

            // Keep the current data
            const currentRecords = worksheet.records.splice(rowIndex);
            const currentData = worksheet.options.data.splice(rowIndex);
            const currentRows = worksheet.rows.splice(rowIndex);

            // Adding lines
            const { rowRecords, rowData, rowNode } = worksheet.formulas.keepChain(() => this.add(rowIndex, numOfRows, rowNumber!, data, currentRows));

            // Copy the data back to the main data
            worksheet.records.push(...currentRecords);
            worksheet.options.data.push(...currentData);
            worksheet.rows.push(...currentRows);

            for (let j = rowIndex; j < worksheet.rows.length; j++) {
                worksheet.rows[j].y = j;
            }

            for (let j = rowIndex; j < worksheet.records.length; j++) {
                for (const cell of worksheet.records[j]) {
                    cell.y = j;
                }
            }

            // Keep history
            worksheet.history.add({
                action: 'insertRow',
                rowNumber,
                numOfRows,
                insertBefore,
                rowRecords,
                rowData,
                rowNode,
                ...(merges && { merges }),
            });

            // The selection and the editor stay on their cells
            worksheet.editor.shift(1, rowIndex, numOfRows);
            worksheet.selection.shift(1, rowIndex, numOfRows);

            // Remove table references
            worksheet.formulas.updateTableReferences({ row: Number(rowNumber), inserted: { axis: 1, index: rowIndex, count: numOfRows } });

            // Respect pagination. The positions of the cells are up to date, for the editor of a row that leaves the page.
            if (worksheet.options.pagination! > 0) {
                worksheet.page(worksheet.view.pageNumber);
            }

            // Events
            worksheet.dispatch('oninsertrow', worksheet, onbeforeinsertrowRecords);
        }
    }

    /**
     * Move a row.
     * @param o - The row to move.
     * @param d - The position it goes to.
     * @param ignoreDom - Move only the data, not the tr.
     */
    move(o: number, d: number, ignoreDom?: boolean) {
        const worksheet = this.worksheet;

        // The application moves the rows of the row source mode
        if (worksheet.rowSource) {
            return;
        }

        // The merged ranges that have the row, or that the row goes into, are removed. The other ranges stay and move with
        // their cells.
        const crossed = [...new Set([...worksheet.merges.isRowMerged(o), ...worksheet.merges.isRowMerged(d, Number(o) > Number(d) ? 1 : 0)])];
        if (crossed.length && !askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
            return false;
        }

        if (worksheet.options.search === true) {
            if (worksheet.view.results && worksheet.view.results.length !== worksheet.rows.length) {
                if (askConfirm(worksheet.parent.config, translate('This action will clear your search results. Are you sure?'))) {
                    worksheet.resetSearch();
                } else {
                    return false;
                }
            }

            worksheet.view.results = null;
        }

        // After the questions, so that a cancel keeps the ranges
        const merges = worksheet.merges.removeCrossed(crossed);

        // Place references in the correct position
        worksheet.rows.splice(d, 0, worksheet.rows.splice(o, 1)[0]);
        worksheet.records.splice(d, 0, worksheet.records.splice(o, 1)[0]);
        worksheet.options.data.splice(d, 0, worksheet.options.data.splice(o, 1)[0]);

        const firstAffectedIndex = Math.min(o, d);
        const lastAffectedIndex = Math.max(o, d);

        for (let j = firstAffectedIndex; j <= lastAffectedIndex; j++) {
            worksheet.rows[j].y = j;
        }

        for (let j = firstAffectedIndex; j <= lastAffectedIndex; j++) {
            for (const cell of worksheet.records[j]) {
                cell.y = j;
            }
        }

        // The arguments can be texts
        const from = Number(o);
        const to = Number(d);

        // The rows that a filter shows keep their place in the results
        if (worksheet.view.results) {
            const shown = worksheet.view.results;
            // The other rows keep their order, so only the moved row needs a new place in the results
            const results = shown
                .filter((row) => row !== from)
                .map((row) => {
                    if (from < to && row > from && row <= to) {
                        return row - 1;
                    }
                    if (from > to && row >= to && row < from) {
                        return row + 1;
                    }
                    return row;
                });
            if (results.length < shown.length) {
                const place = results.findIndex((row) => row > to);
                results.splice(place < 0 ? results.length : place, 0, to);
            }
            worksheet.view.results = results;
        }

        // With pagination, lazy loading or filters, the table body has only some rows. The row goes next to a neighbour
        // that the body shows, and leaves the body when it has none there.
        if (!ignoreDom) {
            const moved = worksheet.rows[to].element;
            const shown = worksheet.view.results || null;
            const k = shown ? shown.indexOf(to) : to;
            const next = k < 0 ? undefined : worksheet.rows[shown ? shown[k + 1] : to + 1]?.element;
            const previous = k < 1 ? undefined : worksheet.rows[shown ? shown[k - 1] : to - 1]?.element;
            if (next?.parentNode === worksheet.tbody) {
                worksheet.tbody.insertBefore(moved, next);
            } else if (previous?.parentNode === worksheet.tbody) {
                worksheet.tbody.insertBefore(moved, previous.nextSibling);
            } else if (moved.parentNode === worksheet.tbody) {
                worksheet.tbody.removeChild(moved);
            }
        }

        // Keeping history of changes
        worksheet.history.add({
            action: 'moveRow',
            oldValue: o,
            newValue: d,
            ...(merges && { merges }),
        });

        // The selection and the editor stay on their cells
        worksheet.editor.move(1, Number(o), Number(d));
        worksheet.selection.move(1, Number(o), Number(d));

        // Update table references
        worksheet.formulas.updateTableReferences({ row: Math.min(Number(o), Number(d)) });

        // The page shows the rows of its positions again. The positions of the cells are up to date, for the editor of a
        // row that leaves the page.
        if (worksheet.options.pagination! > 0) {
            worksheet.view.refreshPage();
        }

        // A row that leaves the table body takes its editor with it. The positions of the cells are up to date.
        worksheet.editor.settle();

        // Events
        worksheet.dispatch('onmoverow', worksheet, parseInt(o), parseInt(d), 1);
    }

    /**
     * Delete rows.
     * @param rowNumber - The first row to delete. Without it, the selected rows, or the last row.
     * @param numOfRows - The number of rows to delete.
     */
    delete(rowNumber?: number, numOfRows?: number): false | undefined {
        const worksheet = this.worksheet;

        // The application changes the rows of the row source mode, and calls resetRows
        if (worksheet.rowSource) {
            return false;
        }

        // Global Configuration
        if (worksheet.options.allowDeleteRow !== false) {
            if (worksheet.options.allowDeletingAllRows === true || worksheet.options.data.length > 1) {
                // Delete row definitions
                if (rowNumber == null) {
                    const number = worksheet.selection.getSelectedRows();

                    if (number.length === 0) {
                        rowNumber = worksheet.options.data.length - 1;
                        numOfRows = 1;
                    } else {
                        rowNumber = number[0];
                        numOfRows = number.length;
                    }
                }

                // Last column
                const lastRow = worksheet.options.data.length - 1;

                if (rowNumber == null || rowNumber > lastRow || rowNumber < 0) {
                    rowNumber = lastRow;
                }

                if (!numOfRows) {
                    numOfRows = 1;
                }

                // Do not delete more than the number of records
                if (rowNumber + numOfRows >= worksheet.options.data.length) {
                    numOfRows = worksheet.options.data.length - rowNumber;
                }

                // Onbeforedeleterow
                const onbeforedeleterowRecords = [];
                for (let i = 0; i < numOfRows; i++) {
                    onbeforedeleterowRecords.push(i + rowNumber);
                }

                if (worksheet.dispatch('onbeforedeleterow', worksheet, onbeforedeleterowRecords) === false) {
                    return false;
                }

                if (parseInt(rowNumber) > -1) {
                    // If delete all rows, and set allowDeletingAllRows false, will stay one row. The merge check below needs
                    // the rows that go.
                    const keepsLastRow = worksheet.options.allowDeletingAllRows !== true && lastRow + 1 === numOfRows;
                    if (keepsLastRow) {
                        numOfRows--;
                    }

                    // The merged ranges that have a cell in the deleted rows are removed. The other ranges stay and move with their cells.
                    const crossed = worksheet.merges.inRows(rowNumber, rowNumber + numOfRows - 1);
                    if (
                        crossed.length &&
                        !askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))
                    ) {
                        return false;
                    }

                    // Clear any search
                    if (worksheet.options.search === true) {
                        if (worksheet.view.results && worksheet.view.results.length !== worksheet.rows.length) {
                            if (askConfirm(worksheet.parent.config, translate('This action will clear your search results. Are you sure?'))) {
                                worksheet.resetSearch();
                            } else {
                                return false;
                            }
                        }

                        worksheet.view.results = null;
                    }

                    // After the questions, so that a cancel keeps the ranges
                    const merges = worksheet.merges.removeCrossed(crossed);

                    // After the questions, so that a cancel logs nothing
                    if (keepsLastRow) {
                        console.error('Jspreadsheet: It is not possible to delete the last row');
                    }

                    // An editor in a deleted row closes without a save
                    worksheet.editor.closeIn(1, rowNumber, rowNumber + numOfRows - 1);

                    // Remove node
                    for (let row = rowNumber; row < rowNumber + numOfRows; row++) {
                        if (worksheet.rows[row].element.parentNode === worksheet.tbody) {
                            worksheet.rows[row].element.className = '';
                            worksheet.rows[row].element.parentNode!.removeChild(worksheet.rows[row].element);
                        }
                    }

                    // Remove data
                    const rowRecords = worksheet.records.splice(rowNumber, numOfRows);
                    const rowData = worksheet.options.data.splice(rowNumber, numOfRows);
                    const rowNode = worksheet.rows.splice(rowNumber, numOfRows);

                    for (let j = rowNumber; j < worksheet.rows.length; j++) {
                        worksheet.rows[j].y = j;
                    }

                    for (let j = rowNumber; j < worksheet.records.length; j++) {
                        for (const cell of worksheet.records[j]) {
                            cell.y = j;
                        }
                    }

                    // Keep history
                    const historyRecord: WorksheetHistoryRecord = {
                        action: 'deleteRow',
                        rowNumber,
                        numOfRows,
                        insertBefore: 1,
                        rowRecords,
                        rowData,
                        rowNode,
                        ...(merges && { merges }),
                    };
                    worksheet.history.add(historyRecord);

                    // The selection and the editor stay on their cells. A selection with a deleted cell is removed.
                    worksheet.editor.shift(1, rowNumber, -numOfRows);
                    worksheet.selection.shift(1, rowNumber, -numOfRows);

                    // Remove table references. The references to the deleted cells become #REF!, and the undo gives the
                    // formulas their texts again.
                    const formulas = worksheet.formulas.updateTableReferences({
                        row: Number(rowNumber),
                        removed: { axis: 1, index: Number(rowNumber), count: numOfRows, refErrors: true },
                    });
                    if (formulas.length) {
                        historyRecord.formulas = formulas;
                    }

                    // The page shows the rows of its positions again, also after a delete on an earlier page. The positions
                    // of the cells are up to date, for the editor of a row that leaves the page.
                    if (worksheet.options.pagination! > 0) {
                        worksheet.view.refreshPage();
                    }

                    // Events
                    worksheet.dispatch('ondeleterow', worksheet, onbeforedeleterowRecords);
                }
            } else {
                console.error('Jspreadsheet: It is not possible to delete the last row');
            }
        }
    }

    /**
     * Get the row height
     *
     * @param row - row number (first row is: 0)
     * @return height - current row height
     */
    getHeight(row?: number | string | HTMLElement): string[] | string {
        const worksheet = this.worksheet;

        let data;

        // The row source mode keeps the heights in its layout
        if (worksheet.rowSource) {
            const layout = worksheet.rowSource.layout;
            if (typeof row === 'undefined') {
                const heights: string[] = [];
                for (const run of layout.runs) {
                    if (run.height! > 0) {
                        for (let y = run.first; y <= run.last; y++) {
                            heights[y] = `${run.height}px`;
                        }
                    }
                }
                return heights;
            }
            const y = typeof row === 'object' ? parseInt(row.getAttribute('data-y')!) : parseInt(row as string);
            return `${layout.shownHeight(y)}px`;
        }

        if (typeof row === 'undefined') {
            // Get height of all rows
            data = [];
            for (let j = 0; j < worksheet.rows.length; j++) {
                const h = worksheet.rows[j].element.style.height;
                if (h) {
                    data[j] = h;
                }
            }
        } else {
            // In case the row is an object
            if (typeof row === 'object') {
                row = row.getAttribute('data-y')!;
            }

            // data-y is the row number as text, which works as an index
            data = worksheet.rows[row as number].element.style.height;
        }

        return data;
    }

    /**
     * Set the row height
     *
     * @param row - row number (first row is: 0)
     * @param height - new row height
     * @param oldHeight - old row height
     */
    // oldHeight is the height attribute as text, or the measured height
    setHeight(row: number, height: number, oldHeight?: number) {
        const worksheet = this.worksheet;

        // The row source mode sets the height in its layout
        if (height > 0 && worksheet.rowSource) {
            height = parseInt(height);
            if (!oldHeight) {
                oldHeight = worksheet.rowSource.layout.shownHeight(parseInt(row));
            }
            worksheet.rowSource.setLayout([row], { height });
            worksheet.history.add({ action: 'setHeight', row, oldValue: oldHeight, newValue: height });
            worksheet.dispatch('onresizerow', worksheet, row, height, oldHeight);
            worksheet.selection.updateCorner();
            return;
        }

        if (height > 0) {
            // Oldwidth
            if (!oldHeight) {
                // A resize with the mouse leaves the height attribute
                oldHeight = parseInt(worksheet.rows[row].element.getAttribute('height'));

                if (!oldHeight) {
                    const rect = worksheet.rows[row].element.getBoundingClientRect();
                    oldHeight = toGrid(worksheet, rect.height);
                }
            }

            // Integer
            height = parseInt(height);

            // Set width
            worksheet.rows[row].element.style.height = `${height}px`;

            if (!worksheet.options.rows) {
                worksheet.options.rows = [];
            }

            // Keep options updated
            if (!worksheet.options.rows[row]) {
                worksheet.options.rows[row] = {};
            }
            worksheet.options.rows[row].height = height;

            // Keeping history of changes
            worksheet.history.add({
                action: 'setHeight',
                row,
                oldValue: oldHeight,
                newValue: height,
            });

            // On resize column
            worksheet.dispatch('onresizerow', worksheet, row, height, oldHeight);

            // Update corner position
            worksheet.selection.updateCorner();
        }
    }

    /**
     * Show row
     */
    show(rowNumber: number | number[]) {
        const worksheet = this.worksheet;

        if (!Array.isArray(rowNumber)) {
            rowNumber = [rowNumber];
        }

        // The row source mode keeps the hidden rows in its layout
        if (worksheet.rowSource) {
            worksheet.rowSource.setLayout(rowNumber, { hidden: false });
            return;
        }

        rowNumber.forEach((rowIndex) => {
            worksheet.rows[rowIndex].element.style.display = '';
        });

        // The rows after a frozen row move
        if (rowNumber.some((rowIndex) => rowIndex < this.getFreezeCount())) {
            this.schedulePlace();
        }
    }

    /**
     * Hide row
     */
    hide(rowNumber: number | number[]) {
        const worksheet = this.worksheet;

        if (!Array.isArray(rowNumber)) {
            rowNumber = [rowNumber];
        }

        if (worksheet.rowSource) {
            worksheet.rowSource.setLayout(rowNumber, { hidden: true });
            worksheet.selection.updateTabStop();
            worksheet.editor.settle();
            return;
        }

        rowNumber.forEach((rowIndex) => {
            worksheet.rows[rowIndex].element.style.display = 'none';
        });

        // The rows after a frozen row move
        if (rowNumber.some((rowIndex) => rowIndex < this.getFreezeCount())) {
            this.schedulePlace();
        }

        // The tab stop of the grid can be in a hidden row, and the editor of a hidden row detaches
        worksheet.selection.updateTabStop();
        worksheet.editor.settle();
    }

    /**
     * Get a row data by rowNumber
     */
    getData(rowNumber: number, processed?: boolean): jspreadsheet.CellValue[] | undefined {
        const worksheet = this.worksheet;

        if (processed) {
            return worksheet.records[rowNumber].map((record) => {
                return record.element.innerHTML;
            });
        } else {
            return worksheet.options.data[rowNumber];
        }
    }

    /**
     * Set a row data by rowNumber
     */
    setData(rowNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
        const worksheet = this.worksheet;

        for (let i = 0; i < worksheet.headers.length; i++) {
            // Update cell
            const columnName = getColumnNameFromId([i, rowNumber]);
            // Set value
            if (data[i] != null) {
                worksheet.setValue(columnName, data[i]!, force);
            }
        }
    }

    /**
     * Show index column
     */
    showIndex() {
        const worksheet = this.worksheet;

        worksheet.table.classList.remove('jss_hidden_index');

        // The position of the frozen columns depends on the width of the index column
        if (worksheet.options.freezeColumns! > 0) {
            worksheet.columns.updateFreezePosition();
        }
    }

    /**
     * Hide index column
     */
    hideIndex() {
        const worksheet = this.worksheet;

        worksheet.table.classList.add('jss_hidden_index');

        // The position of the frozen columns depends on the width of the index column
        if (worksheet.options.freezeColumns! > 0) {
            worksheet.columns.updateFreezePosition();
        }
    }
}
