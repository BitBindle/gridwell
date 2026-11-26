import { getCellNameFromCoords } from '../utils/helpers.js';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import libraryBase from '../utils/libraryBase.js';
import type { WorksheetState, Worksheet, SelectedWorksheet } from '../types/internal';

/** The column and the row of a corner of the selection, then of the other corner: [x1, y1, x2, y2]. */
type Range = [number, number, number, number];

/**
 * The focus is on no element, for example after the focused element was removed
 */
function isFocusLost() {
    const active = document.activeElement;

    return !active || active === document.body;
}

/** The corner of the selection that moves: 0 for the cell in selection.range[0] and [1], 1 for the cell in [2] and [3]. */
type Corner = 0 | 1;

/** 0 moves to the last visible cell in that direction (with Ctrl), 1 moves to the next visible cell. */
type Step = 0 | 1;

/**
 * Whether the worksheet has a selection
 */
function hasSelection(worksheet: Worksheet): worksheet is SelectedWorksheet {
    return worksheet.selection.range ? true : false;
}

/**
 * The selection of a worksheet: the selected range, the highlighted cells, the cells that a drag of the corner fills,
 * and the cell that has the keyboard focus. It selects cells and updates their classes, the ARIA attributes, the corner
 * and the toolbar, keeps the selected cell the tab stop of the grid, and tells which cells are selected.
 *
 * Each worksheet has one, as worksheet.selection. updateSelectionFromCoords, resetSelection, selectAll, getSelection,
 * getSelected, getSelectedColumns, getSelectedRows, getRange, isSelected, getHighlighted, up, down, left, right, first and last call it.
 */
export class Selection {
    readonly worksheet: Worksheet;

    /**
     * The selected range, or null when nothing is selected. The column [0] and row [1] of the cell where the selection
     * starts, and the column [2] and row [3] of the cell where it ends. The end can be left of or above the start.
     */
    range: Range | null = null;

    /** The selected range from left to right and top to bottom, or null before the first selection. */
    bounds: Range | null = null;

    /** The records of the selected cells. */
    highlighted: WorksheetState['records'][number] = [];

    /** The cells that a drag of the selection corner fills, while the drag goes on. */
    fill: HTMLTableCellElement[] = [];

    /** The cell that has the roving tabindex (the tab stop of the grid), or null before the grid has one. */
    activeCell: HTMLTableCellElement | null = null;

    /** Whether the keyboard focus is in the grid or in its editor. */
    focusInGrid = false;

    /** The column whose header a press selected first, while the selection is of whole columns, or null. */
    column: number | null = null;

    /** The row whose number a press selected first, while the selection is of whole rows, or null. */
    row: number | null = null;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /** Remove the fill border from the cells that a drag of the corner was going to fill, and forget them. */
    clearFill() {
        // Remove current selection
        for (const cell of this.fill) {
            cell.classList.remove('selection');
            cell.classList.remove('selection-left');
            cell.classList.remove('selection-right');
            cell.classList.remove('selection-top');
            cell.classList.remove('selection-bottom');
        }

        this.fill = [];
    }

    /**
     * Get the number of columns of the merged range that starts at x, y.
     * The colspan attribute counts only the columns that are not hidden.
     */
    private getMergedColspan(element: HTMLElement, x: number, y: number) {
        const worksheet = this.worksheet;

        const merged = worksheet.options.mergeCells?.[getColumnNameFromId([x, y])];

        return merged ? parseInt(merged[0]) : parseInt(element.getAttribute('colspan'));
    }

    updateCorner() {
        const worksheet = this.worksheet;

        // If any selected cells
        if (!worksheet.selection.highlighted?.length) {
            worksheet.corner.style.top = '-2000px';
            worksheet.corner.style.left = '-2000px';
        } else {
            // Get last cell
            const last = worksheet.selection.highlighted[worksheet.selection.highlighted.length - 1].element;
            const lastX = parseInt(last.getAttribute('data-x'));

            const contentRect = worksheet.content.getBoundingClientRect();
            const x1 = contentRect.left;
            const y1 = contentRect.top;

            const lastRect = last.getBoundingClientRect();
            const x2 = lastRect.left;
            const y2 = lastRect.top;
            const w2 = lastRect.width;
            const h2 = lastRect.height;

            const x = x2 - x1 + worksheet.content.scrollLeft + w2 - 4;
            const y = y2 - y1 + worksheet.content.scrollTop + h2 - 4;

            // Place the corner in the correct place
            worksheet.corner.style.top = `${y}px`;
            worksheet.corner.style.left = `${x}px`;

            if (worksheet.options.freezeColumns) {
                const width = worksheet.columns.getFreezeWidth();
                // Only check if the last column is not part of the merged cells
                if (lastX > worksheet.options.freezeColumns - 1 && x2 - x1 + w2 < width) {
                    worksheet.corner.style.display = 'none';
                } else {
                    if (worksheet.options.selectionCopy !== false) {
                        worksheet.corner.style.display = '';
                    }
                }
            } else {
                if (worksheet.options.selectionCopy !== false) {
                    worksheet.corner.style.display = '';
                }
            }
        }

        worksheet.parent.toolbarManager.update(worksheet);
    }

    reset(blur?: boolean) {
        const worksheet = this.worksheet;

        let previousStatus: 0 | 1;

        // Remove style
        if (!worksheet.selection.highlighted?.length) {
            previousStatus = 0;
        } else {
            previousStatus = 1;

            for (const cell of worksheet.selection.highlighted) {
                cell.element.classList.remove('highlight');
                cell.element.classList.remove('highlight-left');
                cell.element.classList.remove('highlight-right');
                cell.element.classList.remove('highlight-top');
                cell.element.classList.remove('highlight-bottom');
                cell.element.classList.remove('highlight-selected');
                cell.element.removeAttribute('aria-selected');

                const px = parseInt(cell.element.getAttribute('data-x'));
                const py = parseInt(cell.element.getAttribute('data-y'));

                // Check for merged cells
                let ux, uy;

                if (cell.element.getAttribute('data-merged')) {
                    const colspan = this.getMergedColspan(cell.element, px, py);
                    const rowspan = parseInt(cell.element.getAttribute('rowspan'));
                    ux = colspan > 0 ? px + (colspan - 1) : px;
                    uy = rowspan > 0 ? py + (rowspan - 1) : py;
                } else {
                    ux = px;
                    uy = py;
                }

                // Remove selected from headers
                for (let j = px; j <= ux; j++) {
                    if (worksheet.headers[j]) {
                        worksheet.headers[j].classList.remove('selected');
                    }
                }

                // Remove selected from rows
                for (let j = py; j <= uy; j++) {
                    if (worksheet.rows[j]) {
                        worksheet.rows[j].element.classList.remove('selected');
                    }
                }
            }
        }

        // Reset highlighted cells
        worksheet.selection.highlighted = [];

        // The tab stop of the grid stays on the active cell, which is not selected now
        if (worksheet.selection.activeCell) {
            worksheet.selection.activeCell.setAttribute('aria-selected', 'false');
        }

        // Reset
        worksheet.selection.range = null;

        // Hide corner
        worksheet.corner.style.top = '-2000px';
        worksheet.corner.style.left = '-2000px';

        if (blur === true && previousStatus === 1) {
            worksheet.dispatch('onblur', worksheet);
        }

        return previousStatus;
    }

    /**
     * Update selection based on two cells
     */
    update(el1: HTMLElement, el2?: HTMLElement, origin?: Event) {
        const worksheet = this.worksheet;

        // Both elements are cells, which have data-x and data-y
        const x1 = parseInt(el1.getAttribute('data-x'));
        const y1 = parseInt(el1.getAttribute('data-y'));

        let x2, y2;
        if (el2) {
            x2 = parseInt(el2.getAttribute('data-x'));
            y2 = parseInt(el2.getAttribute('data-y'));
        } else {
            x2 = x1;
            y2 = y1;
        }

        this.updateFromCoords(x1, y1, x2, y2, origin);
    }

    // A null x1 selects the rows from y1 to y2, and a null y1 selects the columns from x1 to x2
    updateFromCoords(x1: number | null, y1: number | null, x2?: number | null, y2?: number | null, origin?: Event) {
        const worksheet = this.worksheet;

        // select column
        if (y1 == null) {
            y1 = 0;
            y2 = worksheet.rows.length - 1;

            if (x1 == null) {
                return;
            }
        } else if (x1 == null) {
            // select row
            x1 = 0;
            x2 = worksheet.options.data[0].length - 1;
        }

        // Same element
        if (x2 == null) {
            x2 = x1;
        }
        if (y2 == null) {
            y2 = y1;
        }

        // The mouse events pass the text of data-x and data-y. Keep the selection as numbers.
        x1 = parseInt(x1);
        y1 = parseInt(y1);
        x2 = parseInt(x2);
        y2 = parseInt(y2);

        // Selection must be within the existing data
        if (x1 >= worksheet.headers.length) {
            x1 = worksheet.headers.length - 1;
        }
        if (y1 >= worksheet.rows.length) {
            y1 = worksheet.rows.length - 1;
        }
        if (x2 >= worksheet.headers.length) {
            x2 = worksheet.headers.length - 1;
        }
        if (y2 >= worksheet.rows.length) {
            y2 = worksheet.rows.length - 1;
        }

        // Limits
        let borderLeft = null;
        let borderRight = null;
        let borderTop = null;
        let borderBottom = null;

        // Origin & Destination
        let px: number, ux: number;

        if (parseInt(x1) < parseInt(x2)) {
            px = parseInt(x1);
            ux = parseInt(x2);
        } else {
            px = parseInt(x2);
            ux = parseInt(x1);
        }

        let py: number, uy: number;

        if (parseInt(y1) < parseInt(y2)) {
            py = parseInt(y1);
            uy = parseInt(y2);
        } else {
            py = parseInt(y2);
            uy = parseInt(y1);
        }

        // Verify merged columns
        for (let i = px; i <= ux; i++) {
            for (let j = py; j <= uy; j++) {
                if (worksheet.records[j][i]?.element.getAttribute('data-merged')) {
                    const x = parseInt(worksheet.records[j][i].element.getAttribute('data-x'));
                    const y = parseInt(worksheet.records[j][i].element.getAttribute('data-y'));
                    const colspan = this.getMergedColspan(worksheet.records[j][i].element, x, y);
                    const rowspan = parseInt(worksheet.records[j][i].element.getAttribute('rowspan'));

                    if (colspan > 1) {
                        if (x < px) {
                            px = x;
                        }
                        if (x + colspan > ux) {
                            ux = x + colspan - 1;
                        }
                    }

                    if (rowspan) {
                        if (y < py) {
                            py = y;
                        }
                        if (y + rowspan > uy) {
                            uy = y + rowspan - 1;
                        }
                    }
                }
            }
        }

        // Vertical limits
        for (let j = py; j <= uy; j++) {
            if (worksheet.rows[j].element.style.display !== 'none') {
                if (borderTop == null) {
                    borderTop = j;
                }
                borderBottom = j;
            }
        }

        for (let i = px; i <= ux; i++) {
            for (let j = py; j <= uy; j++) {
                // Horizontal limits
                if (worksheet.options.columns?.[i]?.type !== 'hidden') {
                    if (borderLeft == null) {
                        borderLeft = i;
                    }
                    borderRight = i;
                }
            }
        }

        // Create borders
        if (!borderLeft) {
            borderLeft = 0;
        }
        if (!borderRight) {
            borderRight = 0;
        }

        // The top and bottom are null when every selected row is hidden (issue #59)
        const ret = worksheet.dispatch('onbeforeselection', worksheet, borderLeft, borderTop!, borderRight, borderBottom!, origin);
        if (ret === false) {
            return false;
        }

        // Reset Selection
        const previousState = worksheet.resetSelection();

        // Keep selected cell
        worksheet.selection.range = [x1, y1, x2, y2];

        // Add selected cell
        if (worksheet.records[y1][x1]) {
            worksheet.records[y1][x1].element.classList.add('highlight-selected');
        }

        // Redefining styles
        for (let i = px; i <= ux; i++) {
            for (let j = py; j <= uy; j++) {
                if (worksheet.rows[j].element.style.display !== 'none' && worksheet.records[j][i].element.style.display !== 'none') {
                    worksheet.records[j][i].element.classList.add('highlight');
                    worksheet.records[j][i].element.setAttribute('aria-selected', 'true');
                    worksheet.selection.highlighted.push(worksheet.records[j][i]);
                }
            }
        }

        for (let i = borderLeft; i <= borderRight; i++) {
            if (
                worksheet.options.columns?.[i]?.type !== 'hidden' &&
                worksheet.cols[i].colElement.style &&
                worksheet.cols[i].colElement.style.display !== 'none'
            ) {
                // Top border. The borders are null when each row of the range is hidden.
                if (worksheet.records[borderTop!]?.[i]) {
                    worksheet.records[borderTop!][i].element.classList.add('highlight-top');
                }
                // Bottom border
                if (worksheet.records[borderBottom!]?.[i]) {
                    worksheet.records[borderBottom!][i].element.classList.add('highlight-bottom');
                }
                // Add selected from headers
                worksheet.headers[i].classList.add('selected');
            }
        }

        for (let j = borderTop!; j <= borderBottom!; j++) {
            if (worksheet.rows[j] && worksheet.rows[j].element.style.display !== 'none') {
                // Left border
                worksheet.records[j][borderLeft].element.classList.add('highlight-left');
                // Right border
                worksheet.records[j][borderRight].element.classList.add('highlight-right');
                // Add selected from rows
                worksheet.rows[j].element.classList.add('selected');
            }
        }

        worksheet.selection.bounds = [borderLeft, borderTop!, borderRight, borderBottom!];

        // The selected cell is the tab stop of the grid
        this.updateActiveCell(x1, y1, origin);

        // Handle events
        if (previousState === 0) {
            worksheet.dispatch('onfocus', worksheet);

            worksheet.clipboard.clearBorder();
        }

        // The top and bottom are null when every selected row is hidden (issue #59)
        worksheet.dispatch('onselection', worksheet, borderLeft, borderTop!, borderRight, borderBottom!, origin);

        // Find corner cell
        this.updateCorner();
    }

    /**
     * Get the selected columns.
     * @param visibleOnly - Leave out the hidden columns.
     */
    getSelectedColumns(visibleOnly?: boolean) {
        const worksheet = this.worksheet;

        if (!worksheet.selection.range) {
            return [];
        }

        const result = [];

        for (
            let i = Math.min(worksheet.selection.range[0], worksheet.selection.range[2]);
            i <= Math.max(worksheet.selection.range[0], worksheet.selection.range[2]);
            i++
        ) {
            if (!visibleOnly || worksheet.headers[i].style.display !== 'none') {
                result.push(i);
            }
        }

        return result;
    }

    /**
     * Refresh current selection
     */
    refresh() {
        const worksheet = this.worksheet;

        if (worksheet.selection.range) {
            worksheet.updateSelectionFromCoords(
                worksheet.selection.range[0],
                worksheet.selection.range[1],
                worksheet.selection.range[2],
                worksheet.selection.range[3]
            );
        }
    }

    /**
     * Move coords to A1 in case overlaps with an excluded cell
     */
    // type is 1 for rows and 0 for columns
    conditionalUpdate(type: 0 | 1, o: number, d: number) {
        const worksheet = this.worksheet;

        if (type === 1) {
            if (
                worksheet.selection.range &&
                ((o >= worksheet.selection.range[1] && o <= worksheet.selection.range[3]) ||
                    (d >= worksheet.selection.range[1] && d <= worksheet.selection.range[3]))
            ) {
                worksheet.resetSelection();
                return;
            }
        } else {
            if (
                worksheet.selection.range &&
                ((o >= worksheet.selection.range[0] && o <= worksheet.selection.range[2]) ||
                    (d >= worksheet.selection.range[0] && d <= worksheet.selection.range[2]))
            ) {
                worksheet.resetSelection();
                return;
            }
        }
    }

    /**
     * Get the selected rows.
     * @param visibleOnly - Leave out the hidden rows.
     */
    getSelectedRows(visibleOnly?: boolean) {
        const worksheet = this.worksheet;

        if (!worksheet.selection.range) {
            return [];
        }

        const result = [];

        for (
            let i = Math.min(worksheet.selection.range[1], worksheet.selection.range[3]);
            i <= Math.max(worksheet.selection.range[1], worksheet.selection.range[3]);
            i++
        ) {
            if (!visibleOnly || worksheet.rows[i].element.style.display !== 'none') {
                result.push(i);
            }
        }

        return result;
    }

    selectAll() {
        const worksheet = this.worksheet;

        // updateSelectionFromCoords sets the range, unless onbeforeselection cancels the selection
        worksheet.updateSelectionFromCoords(0, 0, worksheet.headers.length - 1, worksheet.records.length - 1);
    }

    get(): [number, number, number, number] | null {
        const worksheet = this.worksheet;

        if (!worksheet.selection.range) {
            return null;
        }

        return [
            Math.min(worksheet.selection.range[0], worksheet.selection.range[2]),
            Math.min(worksheet.selection.range[1], worksheet.selection.range[3]),
            Math.max(worksheet.selection.range[0], worksheet.selection.range[2]),
            Math.max(worksheet.selection.range[1], worksheet.selection.range[3]),
        ];
    }

    // The names of the cells, or their records. The public typings have an overload for each.
    getSelected(columnNameOnly?: boolean): (string | Worksheet['records'][number][number])[] {
        const worksheet = this.worksheet;

        const selectedRange = this.get();

        if (!selectedRange) {
            return [];
        }

        const cells = [];

        for (let y = selectedRange[1]; y <= selectedRange[3]; y++) {
            for (let x = selectedRange[0]; x <= selectedRange[2]; x++) {
                if (columnNameOnly) {
                    cells.push(getCellNameFromCoords(x, y));
                } else {
                    cells.push(worksheet.records[y][x]);
                }
            }
        }

        return cells;
    }

    getRange() {
        const worksheet = this.worksheet;

        const selectedRange = this.get();

        if (!selectedRange) {
            return '';
        }

        const start = getCellNameFromCoords(selectedRange[0], selectedRange[1]);
        const end = getCellNameFromCoords(selectedRange[2], selectedRange[3]);

        if (start === end) {
            return `${worksheet.options.worksheetName}!${start}`;
        }

        return `${worksheet.options.worksheetName}!${start}:${end}`;
    }

    isSelected(x: number, y: number) {
        const worksheet = this.worksheet;

        const selection = this.get();

        // No cell is selected
        if (!selection) {
            return false;
        }

        return x >= selection[0] && x <= selection[2] && y >= selection[1] && y <= selection[3];
    }

    getHighlighted(): [number, number, number, number][] {
        const worksheet = this.worksheet;

        const selection = this.get();

        if (selection) {
            return [selection];
        }

        return [];
    }

    /**
     * A cell can be the tab stop when its row is in the table and the cell and its row are not hidden
     */
    private isVisibleCell(td: HTMLTableCellElement | null | undefined) {
        const worksheet = this.worksheet;

        return (
            !!td &&
            !!td.parentNode &&
            td.parentNode.parentNode === worksheet.tbody &&
            td.style.display !== 'none' &&
            (td.parentNode as HTMLElement).style.display !== 'none'
        );
    }

    /**
     * First visible cell of the rows in the table
     */
    private getFirstVisibleCell() {
        const worksheet = this.worksheet;

        for (const child of worksheet.tbody.children) {
            const tr = child as HTMLTableRowElement;
            if (tr.style.display !== 'none') {
                // The first child is the row number
                for (let i = 1; i < tr.children.length; i++) {
                    if ((tr.children[i] as HTMLTableCellElement).style.display !== 'none') {
                        return tr.children[i] as HTMLTableCellElement;
                    }
                }
            }
        }

        return null;
    }

    /**
     * Roving tabindex: the active cell is the only cell with tabindex 0. The grid is multiselectable, so the active cell
     * gives aria-selected="false" when it is not selected.
     */
    private setActiveCell(td: HTMLTableCellElement | null | undefined) {
        const worksheet = this.worksheet;

        const previous = worksheet.selection.activeCell;
        if (previous && previous !== td) {
            previous.tabIndex = -1;
            previous.removeAttribute('aria-label');
            if (previous.getAttribute('aria-selected') === 'false') {
                previous.removeAttribute('aria-selected');
            }
        }

        worksheet.selection.activeCell = td || null;

        if (td) {
            // The editor of a cell takes the focus, so the cell is not a focus target while the editor is open
            if (!worksheet.editor.isOpen() || worksheet.editor.cell !== td) {
                td.tabIndex = 0;
            }
            worksheet.accessibility.setCellLabel(td);
            if (td.getAttribute('aria-selected') !== 'true') {
                td.setAttribute('aria-selected', 'false');
            }
        }
    }

    /**
     * The focused element in the document or in the shadow root of the worksheet
     */
    private getActiveElement() {
        const worksheet = this.worksheet;

        const root = worksheet.table.getRootNode() as Document | ShadowRoot;

        return root.activeElement || null;
    }

    /**
     * The tab stop of the grid is the selected cell. When no cell is selected, or the row of the selected cell is not in
     * the table, for example after lazy loading, pagination, a search or deleteRow, the active cell stays the tab stop if it
     * is still in the table, else the first visible cell is the tab stop. When the focused cell left the table, or the table
     * attached its row again, the focus fell to the body, and the tab stop gets it back. A focused cell that is no longer the
     * tab stop also gives the focus to the tab stop.
     */
    updateTabStop() {
        const worksheet = this.worksheet;

        let td: HTMLTableCellElement | null | undefined = null;
        if (worksheet.selection.range) {
            const x = worksheet.selection.range[0];
            const y = worksheet.selection.range[1];
            td = worksheet.records[y]?.[x] ? worksheet.records[y][x].element : null;
        }

        if (!this.isVisibleCell(td)) {
            td = this.isVisibleCell(worksheet.selection.activeCell) ? worksheet.selection.activeCell : this.getFirstVisibleCell();
        }

        this.setActiveCell(td);

        if (td && !worksheet.editor.isOpen() && libraryBase.jspreadsheet.current === worksheet) {
            const active = this.getActiveElement();
            // The focus was in the grid at the last selection
            const lost = worksheet.selection.focusInGrid === true && isFocusLost();
            const otherCell = !!active && active !== td && active.tagName === 'TD' && worksheet.tbody.contains(active);
            if (lost || otherCell) {
                td.focus({ preventScroll: true });
            }
        }
    }

    /**
     * When the focus goes from the grid to an element outside the spreadsheet, a later change of the grid does not take it back
     */
    setFocusEvents() {
        const worksheet = this.worksheet;

        worksheet.table.addEventListener('focusout', (e) => {
            if (e.relatedTarget && (e.relatedTarget as Node).nodeType === 1 && !worksheet.parent.element.contains(e.relatedTarget as Node)) {
                worksheet.selection.focusInGrid = false;
            }
        });
    }

    /**
     * The selected cell is the tab stop of the grid. It also gets the focus when the focus is in the grid, after a press on
     * the grid, or when the focus fell from the grid to the body, for example after lazy loading removed the focused row.
     * A selection does not take the focus from other parts of the page.
     */
    updateActiveCell(x: number, y: number, origin: Event | undefined) {
        const worksheet = this.worksheet;

        const td = worksheet.records[y]?.[x]?.element;
        if (!td) {
            return;
        }

        this.setActiveCell(td);

        const active = this.getActiveElement();
        const pointer = origin?.target && (origin.target as Node).nodeType === 1 && worksheet.table.contains(origin.target as Node);
        const inGrid =
            pointer || (!!active && (active === worksheet.parent.element || active === worksheet.clipboard.textarea || worksheet.table.contains(active)));
        const dropped = isFocusLost() && libraryBase.jspreadsheet.current === worksheet;
        // The focus was in the grid at the last selection
        worksheet.selection.focusInGrid = inGrid || (dropped && worksheet.selection.focusInGrid === true);

        if (worksheet.selection.focusInGrid && !worksheet.editor.isOpen() && active !== td) {
            td.focus({ preventScroll: true });
        }
    }

    /**
     * After the editor closes, the cell is a focus target again. It gets the focus back when the focus fell to the body
     * with the removed editor.
     */
    restoreActiveCell(cell: HTMLTableCellElement) {
        const worksheet = this.worksheet;

        if (worksheet.selection.activeCell !== cell) {
            cell.tabIndex = -1;
            return;
        }

        cell.tabIndex = 0;
        worksheet.accessibility.setCellLabel(cell);

        setTimeout(() => {
            if (isFocusLost() && libraryBase.jspreadsheet.current === worksheet && worksheet.selection.activeCell === cell && !worksheet.editor.isOpen()) {
                cell.focus({ preventScroll: true });
            }
        });
    }

    /**
     * Move the selection from a cell that a merged range hides to the first cell of the range
     */
    private toFirstCell() {
        // The moves call it after they check that the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        const element = worksheet.records[worksheet.selection.range[1]][worksheet.selection.range[0]].element;

        if (element.getAttribute('data-merged')) {
            worksheet.selection.range[0] = parseInt(element.getAttribute('data-x'));
            worksheet.selection.range[1] = parseInt(element.getAttribute('data-y'));
        }
    }

    private upGet(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);
        y = parseInt(y);
        for (let j = y - 1; j >= 0; j--) {
            if (worksheet.records[j][x].element.style.display !== 'none' && worksheet.rows[j].element.style.display !== 'none') {
                if (worksheet.records[j][x].element.getAttribute('data-merged')) {
                    if (worksheet.records[j][x].element === worksheet.records[y][x].element) {
                        continue;
                    }
                }
                y = j;
                break;
            }
        }

        return y;
    }

    private upVisible(group: Corner, direction: Step) {
        // The moves call it after they check that the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        let x, y;

        if (group === 0) {
            x = parseInt(worksheet.selection.range[0]);
            y = parseInt(worksheet.selection.range[1]);
        } else {
            x = parseInt(worksheet.selection.range[2]);
            y = parseInt(worksheet.selection.range[3]);
        }

        if (direction === 0) {
            for (let j = 0; j < y; j++) {
                if (worksheet.records[j][x].element.style.display !== 'none' && worksheet.rows[j].element.style.display !== 'none') {
                    y = j;
                    break;
                }
            }
        } else {
            y = this.upGet(x, y);
        }

        if (group === 0) {
            worksheet.selection.range[0] = x;
            worksheet.selection.range[1] = y;
        } else {
            worksheet.selection.range[2] = x;
            worksheet.selection.range[3] = y;
        }
    }

    up(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (worksheet.selection.range[3] > 0) {
                this.upVisible(1, ctrlKey ? 0 : 1);
            }
        } else {
            if (worksheet.selection.range[1] > 0) {
                this.upVisible(0, ctrlKey ? 0 : 1);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        // Update selection
        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );

        // Change page
        if (worksheet.options.lazyLoading === true) {
            if (worksheet.selection.range[1] === 0 || worksheet.selection.range[3] === 0) {
                worksheet.view.loadPage(0);
                worksheet.updateSelectionFromCoords(
                    worksheet.selection.range[0],
                    worksheet.selection.range[1],
                    worksheet.selection.range[2],
                    worksheet.selection.range[3]
                );
            } else {
                if (worksheet.view.loadValidation()) {
                    worksheet.updateSelectionFromCoords(
                        worksheet.selection.range[0],
                        worksheet.selection.range[1],
                        worksheet.selection.range[2],
                        worksheet.selection.range[3]
                    );
                } else {
                    const item = parseInt((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));
                    if (worksheet.selection.range[1] - item < 30) {
                        worksheet.view.loadUp();
                        worksheet.updateSelectionFromCoords(
                            worksheet.selection.range[0],
                            worksheet.selection.range[1],
                            worksheet.selection.range[2],
                            worksheet.selection.range[3]
                        );
                    }
                }
            }
        } else if (worksheet.options.pagination! > 0) {
            const pageNumber = worksheet.whichPage(worksheet.selection.range[3]);
            if (pageNumber !== worksheet.view.pageNumber) {
                worksheet.page(pageNumber);
            }
        }

        this.updateScroll(1);
    }

    rightGet(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);
        y = parseInt(y);

        for (let i = x + 1; i < worksheet.headers.length; i++) {
            if (worksheet.records[y][i].element.style.display !== 'none') {
                if (worksheet.records[y][i].element.getAttribute('data-merged')) {
                    if (worksheet.records[y][i].element === worksheet.records[y][x].element) {
                        continue;
                    }
                }
                x = i;
                break;
            }
        }

        return x;
    }

    private rightVisible(group: Corner, direction: Step) {
        // The moves call it after they check that the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        let x, y;

        if (group === 0) {
            x = parseInt(worksheet.selection.range[0]);
            y = parseInt(worksheet.selection.range[1]);
        } else {
            x = parseInt(worksheet.selection.range[2]);
            y = parseInt(worksheet.selection.range[3]);
        }

        if (direction === 0) {
            for (let i = worksheet.headers.length - 1; i > x; i--) {
                if (worksheet.records[y][i].element.style.display !== 'none') {
                    x = i;
                    break;
                }
            }
        } else {
            x = this.rightGet(x, y);
        }

        if (group === 0) {
            worksheet.selection.range[0] = x;
            worksheet.selection.range[1] = y;
        } else {
            worksheet.selection.range[2] = x;
            worksheet.selection.range[3] = y;
        }
    }

    right(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (worksheet.selection.range[2] < worksheet.headers.length - 1) {
                this.rightVisible(1, ctrlKey ? 0 : 1);
            }
        } else {
            if (worksheet.selection.range[0] < worksheet.headers.length - 1) {
                this.rightVisible(0, ctrlKey ? 0 : 1);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );
        this.updateScroll(2);
    }

    downGet(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);
        y = parseInt(y);
        for (let j = y + 1; j < worksheet.rows.length; j++) {
            if (worksheet.records[j][x].element.style.display !== 'none' && worksheet.rows[j].element.style.display !== 'none') {
                if (worksheet.records[j][x].element.getAttribute('data-merged')) {
                    if (worksheet.records[j][x].element === worksheet.records[y][x].element) {
                        continue;
                    }
                }
                y = j;
                break;
            }
        }

        return y;
    }

    private downVisible(group: Corner, direction: Step) {
        // The moves call it after they check that the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        let x, y;

        if (group === 0) {
            x = parseInt(worksheet.selection.range[0]);
            y = parseInt(worksheet.selection.range[1]);
        } else {
            x = parseInt(worksheet.selection.range[2]);
            y = parseInt(worksheet.selection.range[3]);
        }

        if (direction === 0) {
            for (let j = worksheet.rows.length - 1; j > y; j--) {
                if (worksheet.records[j][x].element.style.display !== 'none' && worksheet.rows[j].element.style.display !== 'none') {
                    y = j;
                    break;
                }
            }
        } else {
            y = this.downGet(x, y);
        }

        if (group === 0) {
            worksheet.selection.range[0] = x;
            worksheet.selection.range[1] = y;
        } else {
            worksheet.selection.range[2] = x;
            worksheet.selection.range[3] = y;
        }
    }

    down(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (worksheet.selection.range[3] < worksheet.records.length - 1) {
                this.downVisible(1, ctrlKey ? 0 : 1);
            }
        } else {
            if (worksheet.selection.range[1] < worksheet.records.length - 1) {
                this.downVisible(0, ctrlKey ? 0 : 1);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );

        // Change page
        if (worksheet.options.lazyLoading === true) {
            if (worksheet.selection.range[1] === worksheet.records.length - 1 || worksheet.selection.range[3] === worksheet.records.length - 1) {
                worksheet.view.loadPage(-1);
                worksheet.updateSelectionFromCoords(
                    worksheet.selection.range[0],
                    worksheet.selection.range[1],
                    worksheet.selection.range[2],
                    worksheet.selection.range[3]
                );
            } else {
                if (worksheet.view.loadValidation()) {
                    worksheet.updateSelectionFromCoords(
                        worksheet.selection.range[0],
                        worksheet.selection.range[1],
                        worksheet.selection.range[2],
                        worksheet.selection.range[3]
                    );
                } else {
                    const item = parseInt((worksheet.tbody.lastChild as HTMLElement).getAttribute('data-y'));
                    if (item - worksheet.selection.range[3] < 30) {
                        worksheet.view.loadDown();
                        worksheet.updateSelectionFromCoords(
                            worksheet.selection.range[0],
                            worksheet.selection.range[1],
                            worksheet.selection.range[2],
                            worksheet.selection.range[3]
                        );
                    }
                }
            }
        } else if (worksheet.options.pagination! > 0) {
            const pageNumber = worksheet.whichPage(worksheet.selection.range[3]);
            if (pageNumber !== worksheet.view.pageNumber) {
                worksheet.page(pageNumber);
            }
        }

        this.updateScroll(3);
    }

    private leftGet(x: number, y: number) {
        const worksheet = this.worksheet;

        x = parseInt(x);
        y = parseInt(y);
        for (let i = x - 1; i >= 0; i--) {
            if (worksheet.records[y][i].element.style.display !== 'none') {
                if (worksheet.records[y][i].element.getAttribute('data-merged')) {
                    if (worksheet.records[y][i].element === worksheet.records[y][x].element) {
                        continue;
                    }
                }
                x = i;
                break;
            }
        }

        return x;
    }

    private leftVisible(group: Corner, direction: Step) {
        // The moves call it after they check that the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        let x, y;

        if (group === 0) {
            x = parseInt(worksheet.selection.range[0]);
            y = parseInt(worksheet.selection.range[1]);
        } else {
            x = parseInt(worksheet.selection.range[2]);
            y = parseInt(worksheet.selection.range[3]);
        }

        if (direction === 0) {
            for (let i = 0; i < x; i++) {
                if (worksheet.records[y][i].element.style.display !== 'none') {
                    x = i;
                    break;
                }
            }
        } else {
            x = this.leftGet(x, y);
        }

        if (group === 0) {
            worksheet.selection.range[0] = x;
            worksheet.selection.range[1] = y;
        } else {
            worksheet.selection.range[2] = x;
            worksheet.selection.range[3] = y;
        }
    }

    left(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (worksheet.selection.range[2] > 0) {
                this.leftVisible(1, ctrlKey ? 0 : 1);
            }
        } else {
            if (worksheet.selection.range[0] > 0) {
                this.leftVisible(0, ctrlKey ? 0 : 1);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );
        this.updateScroll(0);
    }

    first(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (ctrlKey) {
                worksheet.selection.range[3] = 0;
            } else {
                this.leftVisible(1, 0);
            }
        } else {
            if (ctrlKey) {
                worksheet.selection.range[1] = 0;
            } else {
                this.leftVisible(0, 0);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        // Change page
        if (worksheet.options.lazyLoading === true && (worksheet.selection.range[1] === 0 || worksheet.selection.range[3] === 0)) {
            worksheet.view.loadPage(0);
        } else if (worksheet.options.pagination! > 0) {
            const pageNumber = worksheet.whichPage(worksheet.selection.range[3]);
            if (pageNumber !== worksheet.view.pageNumber) {
                worksheet.page(pageNumber);
            }
        }

        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );
        this.updateScroll(1);
    }

    last(shiftKey?: boolean, ctrlKey?: boolean) {
        const worksheet = this.worksheet;

        // Without a selection, there is nothing to move
        if (!hasSelection(worksheet)) {
            return;
        }

        if (shiftKey) {
            if (ctrlKey) {
                worksheet.selection.range[3] = worksheet.records.length - 1;
            } else {
                this.rightVisible(1, 0);
            }
        } else {
            if (ctrlKey) {
                worksheet.selection.range[1] = worksheet.records.length - 1;
            } else {
                this.rightVisible(0, 0);
            }
            this.toFirstCell();
            worksheet.selection.range[2] = worksheet.selection.range[0];
            worksheet.selection.range[3] = worksheet.selection.range[1];
        }

        // Change page
        if (
            worksheet.options.lazyLoading === true &&
            (worksheet.selection.range[1] === worksheet.records.length - 1 || worksheet.selection.range[3] === worksheet.records.length - 1)
        ) {
            worksheet.view.loadPage(-1);
        } else if (worksheet.options.pagination! > 0) {
            const pageNumber = worksheet.whichPage(worksheet.selection.range[3]);
            if (pageNumber !== worksheet.view.pageNumber) {
                worksheet.page(pageNumber);
            }
        }

        worksheet.updateSelectionFromCoords(
            worksheet.selection.range[0],
            worksheet.selection.range[1],
            worksheet.selection.range[2],
            worksheet.selection.range[3]
        );
        this.updateScroll(3);
    }

    /**
     * Update scroll position based on the selection
     */
    private updateScroll(direction: 0 | 1 | 2 | 3) {
        const worksheet = this.worksheet;

        // Jspreadsheet Container information
        const contentRect = worksheet.content.getBoundingClientRect();
        const x1 = contentRect.left;
        const y1 = contentRect.top;
        const w1 = contentRect.width;
        const h1 = contentRect.height;

        // Direction Left or Up
        const reference = worksheet.records[worksheet.selection.range![3]][worksheet.selection.range![2]].element;

        // Reference
        const referenceRect = reference.getBoundingClientRect();
        const x2 = referenceRect.left;
        const y2 = referenceRect.top;
        const w2 = referenceRect.width;
        const h2 = referenceRect.height;

        let x, y;

        // Direction
        if (direction === 0 || direction === 1) {
            x = x2 - x1 + worksheet.content.scrollLeft;
            y = y2 - y1 + worksheet.content.scrollTop - 2;
        } else {
            x = x2 - x1 + worksheet.content.scrollLeft + w2;
            y = y2 - y1 + worksheet.content.scrollTop + h2;
        }

        // Top position check
        if (y > worksheet.content.scrollTop + 30 && y < worksheet.content.scrollTop + h1) {
            // In the viewport
        } else {
            // Out of viewport
            if (y < worksheet.content.scrollTop + 30) {
                worksheet.content.scrollTop = y - h2;
            } else {
                worksheet.content.scrollTop = y - (h1 - 2);
            }
        }

        // Freeze columns?
        const freezed = worksheet.columns.getFreezeWidth();

        // The frozen columns stay at the left of the content only when scrollLeft is more than the width of the index column.
        // Then a cell whose left border is under the frozen columns goes to the right of the frozen columns.
        const indexWidth = worksheet.columns.getIndexWidth();
        const frozen = freezed > 0 && worksheet.content.scrollLeft > indexWidth;
        const left = x2 - x1 + worksheet.content.scrollLeft;
        const underFreezed = frozen && left < worksheet.content.scrollLeft + freezed;

        // Left position check - TODO: change that to the bottom border of the element
        if (
            (frozen && worksheet.selection.range![2] < worksheet.options.freezeColumns!) ||
            (!underFreezed && x > worksheet.content.scrollLeft + freezed && x < worksheet.content.scrollLeft + w1)
        ) {
            // In the viewport
        } else if (underFreezed) {
            worksheet.content.scrollLeft = left - freezed - 1 < indexWidth ? 0 : left - freezed - 1;
        } else {
            // Out of viewport
            if (x < worksheet.content.scrollLeft + 30) {
                worksheet.content.scrollLeft = x;
                if (worksheet.content.scrollLeft < 50) {
                    worksheet.content.scrollLeft = 0;
                }
            } else if (x < worksheet.content.scrollLeft + freezed) {
                worksheet.content.scrollLeft = x - freezed - 1;
            } else {
                worksheet.content.scrollLeft = x - (w1 - 20);
            }
        }
    }
}
