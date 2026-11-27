import { ColumnLayout } from '../utils/columnLayout.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import { toGrid, toScreen, zoomOf } from '../utils/zoom.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/** The columns left and right of the view that the window holds. */
const defaultBuffer = 10;

/** The columns that the view can come near an edge of the window before the window moves. */
const margin = 3;

/**
 * The largest width of the scroll area in screen pixels. Firefox stops sticky positioning at about 8,947,848 px, so the
 * frozen headers would scroll out of the view after it.
 */
const defaultMaxScrollWidth = 8000000;

/** The most columns left of the window that the window adds to show the first cell of a merged range. */
const maxColumnMergeReach = 100;

/** The scroll area: the frozen columns, the view and the columns after the frozen columns, in grid pixels. */
interface Geometry {
    frozen: number;
    frozenWidth: number;
    indexWidth: number;
    view: number;
    total: number;
    scaled: boolean;
    flow: number;
    maxScroll: number;
    maxLeft: number;
}

/** The runs of consecutive columns of a sorted list, as [from, to] pairs. */
function runs(list: number[]) {
    const result: [number, number][] = [];
    for (const x of list) {
        const last = result[result.length - 1];
        if (last && last[1] + 1 === x) {
            last[1] = x;
        } else {
            result.push([x, x]);
        }
    }
    return result;
}

/** A cell or a col that stands for columns without a cell. The class marks it, so code that reads the cells skips it. */
function spacerCell<K extends 'td' | 'col'>(tag: K): HTMLElementTagNameMap[K] {
    const element = document.createElement(tag);
    element.className = 'jss_colspacer';
    if (tag === 'col') {
        element.setAttribute('width', '0');
    } else {
        element.setAttribute('aria-hidden', 'true');
        element.setAttribute('role', 'presentation');
        element.style.cssText = 'padding:0;border:0;background:transparent';
    }
    return element;
}

/** Insert the items of the columns left before leftAnchor, and of the columns right before rightAnchor. */
function insertColumns(parent: Element, left: number[], right: number[], leftAnchor: Element, rightAnchor: Element, item: (x: number) => Element) {
    for (const [list, anchor] of [
        [left, leftAnchor],
        [right, rightAnchor],
    ] as const) {
        if (list.length) {
            const fragment = document.createDocumentFragment();
            for (const x of list) {
                fragment.appendChild(item(x));
            }
            parent.insertBefore(fragment, anchor);
        }
    }
}

/**
 * The column window of the row source mode. Each row has a cell only for the frozen columns and for a window of columns
 * around the view. A spacer cell stands for the columns between the frozen columns and the window, and a second spacer
 * cell for the columns after the window. The header row and the colgroup have the same items. The spacers give the
 * scroll area the width of all columns. Above maxScrollWidth, a scroll position maps to a column in proportion. The
 * arrays headers and cols keep an item for each column, but only the items of the frozen columns and of the window are
 * in the table. A row record and its values have items only for these columns.
 */
export class ColumnSource {
    readonly worksheet: Worksheet;

    layout: ColumnLayout;

    /** The first and the last column of the window after the frozen columns. The window is empty when to is below from. */
    from = 0;
    to = -1;

    private readonly buffer: number;
    private readonly maxScrollWidth: number;
    private readonly headerLeft = spacerCell('td');
    private readonly headerRight = spacerCell('td');
    private readonly colLeft = spacerCell('col');
    private readonly colRight = spacerCell('col');
    private expectedLeft: number | null = null;
    private wantedLeft: number | null = null;
    /** The cells that show the text of a merged range that starts left of the window. */
    private readonly moved = new Map<string, number | null>();

    constructor(worksheet: Worksheet, source: jspreadsheet.RowSource) {
        this.worksheet = worksheet;
        this.buffer = source.columnBuffer! > 0 ? Math.floor(source.columnBuffer!) : defaultBuffer;
        this.maxScrollWidth = source.maxScrollWidth! > 0 ? Math.min(source.maxScrollWidth!, defaultMaxScrollWidth) : defaultMaxScrollWidth;
        this.layout = this.makeLayout();
    }

    /** The width of column x from its col: 0 for a hidden column. */
    private columnWidth(x: number) {
        const worksheet = this.worksheet;
        const col = worksheet.cols[x]?.colElement;
        if (!col || col.style.display === 'none' || worksheet.headers[x]?.style.display === 'none') {
            return 0;
        }
        return parseInt(col.getAttribute('width')!) || 0;
    }

    private makeLayout() {
        const widths = new Array<number>(this.worksheet.headers.length);
        for (let x = 0; x < widths.length; x++) {
            widths[x] = this.columnWidth(x);
        }
        return new ColumnLayout(widths);
    }

    frozenCount() {
        return Math.min(this.worksheet.options.freezeColumns || 0, this.worksheet.headers.length);
    }

    /**
     * Make the header row and the colgroup again before the rows, for example after a change of the frozen columns: the
     * frozen columns, a spacer, the window for the scroll position and a second spacer.
     */
    start() {
        const worksheet = this.worksheet;
        this.layout = this.makeLayout();
        this.moved.clear();
        const frozen = this.frozenCount();
        const header = worksheet.headerContainer;
        const colgroup = worksheet.colgroupContainer;
        // The first item is the column of the row numbers. One call removes the other items, because a removal of each item is slow.
        header.replaceChildren(header.firstChild!);
        colgroup.replaceChildren(colgroup.firstChild!);
        for (let x = 0; x < frozen; x++) {
            header.appendChild(worksheet.headers[x]);
            colgroup.appendChild(worksheet.cols[x].colElement);
        }
        header.append(this.headerLeft, this.headerRight);
        colgroup.append(this.colLeft, this.colRight);
        this.from = frozen;
        this.to = frozen - 1;
        this.render(true);
    }

    /** The columns with a cell in each row, in order: the frozen columns and the window. */
    attachedColumns() {
        const result: number[] = [];
        for (let x = 0; x < this.frozenCount(); x++) {
            result.push(x);
        }
        for (let x = this.from; x <= this.to; x++) {
            result.push(x);
        }
        return result;
    }

    /** The columns from first to last with a cell in each row, in order. */
    attachedIn(first: number, last: number) {
        const result: number[] = [];
        const frozen = this.frozenCount();
        for (let x = Math.max(0, first); x <= Math.min(last, frozen - 1); x++) {
            result.push(x);
        }
        for (let x = Math.max(first, this.from); x <= Math.min(last, this.to); x++) {
            result.push(x);
        }
        return result;
    }

    /**
     * The part of the columns first to last that has cells in one block of the table, as [from, to], or null. A range
     * that starts in a frozen column ends at the last frozen column.
     */
    attachedPart(first: number, last: number): [number, number] | null {
        const frozen = this.frozenCount();
        if (first < frozen) {
            return [first, Math.min(last, frozen - 1)];
        }
        if (this.to < this.from || first > this.to || last < this.from) {
            return null;
        }
        return [Math.max(first, this.from), Math.min(last, this.to)];
    }

    /** The number of cells of a row in the table: the row number, the frozen columns, the window and the two spacers. */
    cellCount() {
        return 1 + this.frozenCount() + 2 + Math.max(0, this.to - this.from + 1);
    }

    /** Append the cells of row y to its tr: the frozen columns, a spacer, the window and a second spacer. */
    appendCells(tr: HTMLTableRowElement, append: (x: number) => void) {
        for (let x = 0; x < this.frozenCount(); x++) {
            append(x);
        }
        tr.appendChild(spacerCell('td'));
        for (let x = this.from; x <= this.to; x++) {
            append(x);
        }
        tr.appendChild(spacerCell('td'));
    }

    /** The scroll position of the view in grid pixels. */
    private scrollLeft() {
        return toGrid(this.worksheet, this.worksheet.content.scrollLeft);
    }

    private geometry(): Geometry {
        const worksheet = this.worksheet;
        const layout = this.layout;
        const frozen = this.frozenCount();
        const frozenWidth = layout.offsetOf(frozen);
        const indexWidth = worksheet.columns.getIndexWidth();
        // A view without a width limit would hold all columns, so the view is at most the width of the screen
        const screenWidth = toGrid(worksheet, Math.max(1000, window.innerWidth || 0));
        const view = Math.max(1, Math.min(toGrid(worksheet, worksheet.content.clientWidth), screenWidth) - indexWidth - frozenWidth);
        const total = layout.total;
        // Above 100 %, the limit in grid pixels is lower, so the scroll area stays below the limit of the browser
        const maxWidth = this.maxScrollWidth / Math.max(zoomOf(worksheet), 1);
        const scaled = total > maxWidth;
        const flow = scaled ? maxWidth - frozenWidth : total - frozenWidth;
        const maxScroll = Math.max(1, flow - view);
        const maxLeft = Math.max(0, total - frozenWidth - view);
        return { frozen, frozenWidth, indexWidth, view, total, scaled, flow, maxScroll, maxLeft };
    }

    /**
     * The scroll position in the columns after the frozen columns. The row numbers scroll away first, so the columns
     * start to move after indexWidth pixels. The frozen columns then cover the left part of the view.
     */
    private flowScroll(g: Geometry) {
        const scroll = Math.max(0, this.scrollLeft() - g.indexWidth);
        return g.scaled ? Math.min(scroll, g.maxScroll) : scroll;
    }

    /** The place in the columns of the left edge of the view after the frozen columns. */
    private viewLeft(g: Geometry) {
        const scrollLeft = this.scrollLeft();
        if (this.wantedLeft !== null && this.expectedLeft !== null && Math.abs(scrollLeft - this.expectedLeft) < 2) {
            return this.wantedLeft;
        }
        this.wantedLeft = null;
        this.expectedLeft = null;
        const scroll = this.flowScroll(g);
        return g.frozenWidth + (g.scaled ? (scroll / g.maxScroll) * g.maxLeft : scroll);
    }

    /** The scroll position that shows the place left of the columns at the left edge of the view. */
    private scrollFor(g: Geometry, left: number) {
        const offset = Math.max(0, Math.min(left - g.frozenWidth, g.maxLeft));
        const scroll = g.scaled ? (offset / g.maxLeft) * g.maxScroll : offset;
        return scroll > 0 ? scroll + g.indexWidth : 0;
    }

    private setSpacers(left: number, right: number) {
        this.colLeft.setAttribute('width', String(Math.max(0, Math.round(left))));
        this.colRight.setAttribute('width', String(Math.max(0, Math.round(right))));
        // The spacer rows of the row source mode span all cells of a row
        this.worksheet.rowSource?.setSpacerSpan(this.cellCount());
    }

    /** The start of the window moves left to the first column of each merged range that holds its first column. */
    private mergeReach(from: number, to: number) {
        const source = this.worksheet.rowSource;
        if (!source) {
            return from;
        }
        const frozen = this.frozenCount();
        let result = from;
        for (const y of source.attachedRows()) {
            for (const range of source.mergesIn(y, y)) {
                if (range.y1 === y && range.x1 < result && range.x1 >= frozen && range.x2 >= from && range.x1 <= to && from - range.x1 <= maxColumnMergeReach) {
                    result = range.x1;
                }
            }
        }
        return Math.max(frozen, result);
    }

    /**
     * Render the window of columns for the scroll position. With force, make a new window also when the view is inside
     * the window. oncolumnwindow occurs when the window changes.
     */
    render(force?: boolean) {
        const worksheet = this.worksheet;
        if (!worksheet.content) {
            return;
        }
        const layout = this.layout;
        const count = layout.count;
        const g = this.geometry();
        const scrollBefore = worksheet.content.scrollLeft;
        const left = this.viewLeft(g);
        const first = Math.max(g.frozen, Math.min(count - 1, layout.columnAt(left)));
        const last = Math.max(first, Math.min(count - 1, layout.columnAt(left + g.view)));
        let from = this.from;
        let to = this.to;
        const inside = to >= from && from <= Math.max(g.frozen, first - margin) && to >= Math.min(count - 1, last + margin);
        if (force || !inside) {
            from = Math.max(g.frozen, first - this.buffer);
            to = Math.min(count - 1, last + this.buffer);
            from = this.mergeReach(from, to);
        }
        if (count <= g.frozen) {
            from = g.frozen;
            to = g.frozen - 1;
        }
        // The spacer left of the window puts the column at left at the left edge of the view after the frozen columns
        const scroll = this.flowScroll(g);
        let before = scroll - left + layout.offsetOf(from);
        if (before < 0) {
            // Near the left edge of a scaled scroll area, the columns left of the view do not fit left of it
            const limit = left - scroll;
            let column = layout.columnAt(limit);
            if (layout.offsetOf(column) < limit) {
                column++;
            }
            from = Math.min(Math.max(g.frozen, column), first);
            before = Math.max(0, scroll - left + layout.offsetOf(from));
        }
        let after = g.flow - before - (layout.offsetOf(to + 1) - layout.offsetOf(from));
        if (after < 0) {
            const end = layout.offsetOf(from) + g.flow - before;
            to = Math.max(last, Math.min(to, layout.columnAt(end) - 1));
            after = Math.max(0, g.flow - before - (layout.offsetOf(to + 1) - layout.offsetOf(from)));
        }
        const changed = from !== this.from || to !== this.to;
        if (changed) {
            this.moveWindow(from, to);
        }
        this.setSpacers(before, after);
        // While the window changes its columns, the table can be narrower for a moment, and the browser lowers the scroll
        // position. The spacers are for the scroll position before the change, so it comes back.
        if (changed && worksheet.content.scrollLeft !== scrollBefore) {
            worksheet.content.scrollLeft = scrollBefore;
        }
        if (changed) {
            worksheet.dispatch('oncolumnwindow', worksheet, from, to, first, last);
        }
    }

    /**
     * Make the columns from to to the window: remove the cells, headers and cols of the other columns, and add the missing
     * ones in order. The merged ranges, the selection and the focus follow.
     */
    private moveWindow(from: number, to: number) {
        const worksheet = this.worksheet;
        const source = worksheet.rowSource!;
        const rows = source.attachedRows();
        const oldFrom = this.from;
        const oldTo = this.to;
        source.clearMerges();
        const removed: number[] = [];
        for (let x = oldFrom; x <= oldTo; x++) {
            if (x < from || x > to) {
                removed.push(x);
            }
        }
        // An open editor in a column that leaves the window detaches, as oneditordetach describes
        const editor = worksheet.editor;
        if (editor.isOpen() && editor.x >= oldFrom && editor.x <= oldTo && (editor.x < from || editor.x > to)) {
            if (document.activeElement && editor.cell.contains(document.activeElement)) {
                source.focusTable();
            }
            editor.detach();
        }
        if (removed.length) {
            for (const y of rows) {
                const records = worksheet.records[y];
                const data = worksheet.options.data[y];
                for (const x of removed) {
                    const record = records[x];
                    if (!record) {
                        continue;
                    }
                    source.keepFocus(record.element);
                    record.element.remove();
                    delete records[x];
                    delete data[x];
                }
            }
            for (const x of removed) {
                worksheet.headers[x].remove();
                worksheet.cols[x].colElement.remove();
            }
            if (worksheet.selection.highlighted.length) {
                worksheet.selection.highlighted = worksheet.selection.highlighted.filter((record) => record.element.isConnected);
            }
            this.dropChain(new Set(removed));
        }
        // The columns left of the kept columns go after the left spacer, the others before the right spacer
        const keptFrom = Math.max(from, oldFrom);
        const keptTo = Math.min(to, oldTo);
        const kept = keptFrom <= keptTo;
        const addedLeft: number[] = [];
        const addedRight: number[] = [];
        for (let x = from; x <= to; x++) {
            if (kept && x >= keptFrom && x <= keptTo) {
                continue;
            }
            (kept && x < keptFrom ? addedLeft : addedRight).push(x);
        }
        const added = [...addedLeft, ...addedRight];
        this.from = from;
        this.to = to;
        insertColumns(
            worksheet.headerContainer,
            addedLeft,
            addedRight,
            kept ? worksheet.headers[keptFrom] : this.headerRight,
            this.headerRight,
            (x) => worksheet.headers[x]
        );
        insertColumns(
            worksheet.colgroupContainer,
            addedLeft,
            addedRight,
            kept ? worksheet.cols[keptFrom].colElement : this.colRight,
            this.colRight,
            (x) => worksheet.cols[x].colElement
        );
        if (added.length) {
            for (const y of rows) {
                const row = worksheet.rows[y];
                const given = source.values(y, added);
                const data = worksheet.options.data[y];
                const rightSpacer = row.element.lastElementChild!;
                const anchor = kept ? worksheet.records[y][keptFrom].element : rightSpacer;
                insertColumns(row.element, addedLeft, addedRight, anchor, rightSpacer, (x) => {
                    const value = given?.[x];
                    data[x] = value === undefined || value === null ? '' : value;
                    const element = worksheet.rowManager.createRecord(y, x, data[x]);
                    // A cell of a hidden column is hidden
                    if (worksheet.headers[x].style.display === 'none') {
                        element.style.display = 'none';
                    }
                    return element;
                });
                // With textOverflow, a cell does not overflow into a right neighbour that has a value. A new cell has no
                // right neighbour yet when the neighbour comes later or was attached before it.
                if (worksheet.options.textOverflow === true) {
                    for (const x of added) {
                        const right = worksheet.records[y][x + 1];
                        if (right && (data[x + 1] || right.element.innerHTML)) {
                            worksheet.records[y][x].element.style.overflow = 'hidden';
                        }
                    }
                }
            }
        }
        source.applyMerges();
        if (added.length) {
            const ranges = runs(added.sort((a, b) => a - b));
            for (const y of rows) {
                source.selectRow(y, added);
            }
            source.restoreFocus();
            if (rows.length) {
                worksheet.dispatch('oncolumnsattached', worksheet, ranges, runs(rows));
            }
        }
        source.updateTabStop();
    }

    /** Remove the formula cells of removed columns from the formula chain. Their formulas add them again when they show. */
    private dropChain(columns: Set<number>) {
        const chain = this.worksheet.formulas.chain;
        for (const name of Object.keys(chain)) {
            chain[name] = chain[name].filter((dependent) => !columns.has(getIdFromColumnName(dependent)[0]));
            if (!chain[name].length) {
                delete chain[name];
            }
        }
    }

    /**
     * Show the text of the first cell of a merged range, at (from, y), in the cell (to, y) of the window. The cell takes
     * the wrap of the text, as the first cell has it. The values of the cell do not change.
     */
    moveText(from: number, to: number, y: number) {
        const worksheet = this.worksheet;
        const given = worksheet.rowSource!.values(y, [from]);
        const value = given?.[from] ?? '';
        this.moved.set(getColumnNameFromId([to, y]), from);
        const element = worksheet.records[y][to].element;
        const td = worksheet.cells.create(from, y, value, true);
        element.innerHTML = td.innerHTML;
        element.style.whiteSpace = td.style.whiteSpace;
        // The text goes through the render of the first column, as createRecord does
        const column = worksheet.options.columns?.[from];
        if (column && typeof column.render === 'function') {
            column.render(element, value, from, y, worksheet, column);
        }
    }

    /** The cell (x, y) of the window shows no value, as a cell that a merged range hides. Its value does not change. */
    clearText(x: number, y: number) {
        this.moved.set(getColumnNameFromId([x, y]), null);
        this.worksheet.records[y][x].element.innerHTML = '';
    }

    /**
     * After a change of the cell (x, y), a cell that shows the text of a merged range, or no text, shows it again in place
     * of its own value.
     */
    keepText(x: number, y: number) {
        if (!this.moved.size) {
            return;
        }
        const from = this.moved.get(getColumnNameFromId([x, y]));
        if (from === null) {
            this.clearText(x, y);
        } else if (from !== undefined) {
            this.moveText(from, x, y);
        }
    }

    /**
     * Before a change of the cell (x, y): a cell that shows the text of a merged range, or no text, gets its own content
     * again, for example the input of a checkbox, so that the change finds it. keepText then applies the rule again.
     */
    ownContent(x: number, y: number) {
        if (this.moved.size && this.moved.has(getColumnNameFromId([x, y]))) {
            this.rebuild(x, y);
        }
    }

    /** Make the content and the wrap of the cell (x, y) again from its value, as a new cell has them. */
    private rebuild(x: number, y: number) {
        const worksheet = this.worksheet;
        const element = worksheet.records[y][x].element;
        const td = worksheet.cells.create(x, y, worksheet.options.data[y][x], true);
        element.innerHTML = td.innerHTML;
        element.style.whiteSpace = td.style.whiteSpace;
    }

    /** The cells that showed the text of a merged range show their own value again. */
    restoreMoved() {
        const worksheet = this.worksheet;
        // The map is empty before showValue, so keepText does not apply the text rule again
        const names = [...this.moved.keys()];
        this.moved.clear();
        for (const name of names) {
            const [x, y] = getIdFromColumnName(name);
            if (worksheet.records[y]?.[x]) {
                // The content of the cell can miss the parts of its column type, for example the input of a checkbox
                this.rebuild(x, y);
                worksheet.cells.showValue(x, y);
            }
        }
    }

    /**
     * Scroll so that column x shows, and render its window. A column left of the view goes to the left edge of the view,
     * and a column right of it to the right edge. A column that shows does not move.
     */
    scrollToColumn(x: number) {
        const layout = this.layout;
        x = Math.max(0, Math.min(layout.count - 1, Math.floor(Number(x)) || 0));
        const g = this.geometry();
        if (x < g.frozen) {
            this.render();
            return;
        }
        const left = this.viewLeft(g);
        const columnLeft = layout.offsetOf(x);
        const columnRight = layout.offsetOf(x + 1);
        let wanted = left;
        if (columnLeft < left) {
            wanted = columnLeft;
        } else if (columnRight > left + g.view) {
            wanted = columnRight - columnLeft > g.view ? columnLeft : columnRight - g.view;
        }
        if (wanted !== left) {
            wanted = Math.max(g.frozenWidth, Math.min(wanted, g.frozenWidth + g.maxLeft));
            this.worksheet.content.scrollLeft = toScreen(this.worksheet, this.scrollFor(g, wanted));
            this.expectedLeft = this.scrollLeft();
            this.wantedLeft = wanted;
        }
        this.render();
    }

    /** Read the widths and the hidden columns again, for example after a resize, and render the window. */
    refresh() {
        this.layout = this.makeLayout();
        this.render(true);
    }

    /** The first and the last column of the window after the frozen columns. */
    window(): [number, number] {
        return [this.from, this.to];
    }
}
