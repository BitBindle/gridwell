import { getIdFromColumnName } from '../utils/internalHelpers.js';
import { RowLayout } from '../utils/rowLayout.js';
import { toGrid, toScreen, zoomOf } from '../utils/zoom.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/** The rows above and below the view that the window holds. */
const defaultBuffer = 200;

/** The rows above and below the view that the window holds after a far scroll. */
const farScrollBuffer = 20;

/** The time in milliseconds without a far scroll after which the window gets all rows of the buffer. */
const fillDelay = 100;

/**
 * The largest height of the scroll area in screen pixels. Firefox stops sticky positioning at about 8,947,848 px, so the
 * header row and the frozen rows would scroll out of the view above it.
 */
const defaultMaxScrollHeight = 8000000;

/** The most rows above the window that the window adds to show the first cell of a merged range. */
const maxMergeReach = 1000;

/** The scroll area: the frozen rows, the view and the rows after the frozen rows, in grid pixels. */
interface Geometry {
    frozen: number;
    frozenHeight: number;
    view: number;
    total: number;
    scaled: boolean;
    flow: number;
    maxScroll: number;
    maxTop: number;
}

/** True when the cell and its row are not hidden. */
function shows(cell: HTMLElement) {
    return cell.style.display !== 'none' && (cell.parentNode as HTMLElement | null)?.style.display !== 'none';
}

/** The runs of consecutive rows of a sorted list of rows, as [from, to] pairs. */
function ranges(rows: number[]) {
    const result: [number, number][] = [];
    for (const y of rows) {
        const last = result[result.length - 1];
        if (last && last[1] + 1 === y) {
            last[1] = y;
        } else {
            result.push([y, y]);
        }
    }
    return result;
}

/**
 * Why a merged range cannot be one of the ranges of the sheet, or null when it can: the name must be a cell name, the
 * spans two positive whole numbers for more than one cell, and the range inside the sheet.
 */
export function mergeProblem(name: string, span: readonly unknown[], columns: number, rows: number) {
    if (!/^[A-Z]+[1-9][0-9]*$/.test(name) || !Array.isArray(span) || span.length !== 2) {
        return 'Invalid merged properties';
    }
    const [colspan, rowspan] = span as number[];
    if (!Number.isInteger(colspan) || !Number.isInteger(rowspan) || colspan < 1 || rowspan < 1 || colspan * rowspan < 2) {
        return 'Invalid merged properties';
    }
    const [x, y] = getIdFromColumnName(name);
    if (x + colspan > columns || y + rowspan > rows) {
        return 'Invalid merged properties';
    }
    return null;
}

/**
 * The cells that the accepted merged ranges cover, by column. The start rows of all candidate ranges are known first, so
 * each column keeps them sorted, with a Fenwick tree of the largest last row of the accepted ranges up to each place. A
 * check and an add then take O(log n) for each column of a range, in any order of the keys.
 */
class CoveredColumns {
    private readonly columns = new Map<number, { starts: number[]; ends: Float64Array }>();

    constructor(candidates: Iterable<readonly [number, number, number]>) {
        const starts = new Map<number, number[]>();
        for (const [x1, x2, y1] of candidates) {
            for (let x = x1; x <= x2; x++) {
                let list = starts.get(x);
                if (!list) {
                    list = [];
                    starts.set(x, list);
                }
                list.push(y1);
            }
        }
        for (const [x, list] of starts) {
            list.sort((a, b) => a - b);
            const unique = list.filter((y, i) => i === 0 || y !== list[i - 1]);
            this.columns.set(x, { starts: unique, ends: new Float64Array(unique.length + 1).fill(-1) });
        }
    }

    /** The number of start rows of the column that are at or before row y. */
    private static countUpTo(starts: number[], y: number) {
        let low = 0;
        let high = starts.length;
        while (low < high) {
            const middle = (low + high) >> 1;
            if (starts[middle] <= y) {
                low = middle + 1;
            } else {
                high = middle;
            }
        }
        return low;
    }

    /** True when an accepted range covers a cell of the columns x1 to x2 and the rows y1 to y2. */
    crosses(x1: number, x2: number, y1: number, y2: number) {
        for (let x = x1; x <= x2; x++) {
            const column = this.columns.get(x);
            if (!column) {
                continue;
            }
            // The largest last row of the accepted ranges that start at or before y2
            let last = -1;
            for (let i = CoveredColumns.countUpTo(column.starts, y2); i > 0; i -= i & -i) {
                last = Math.max(last, column.ends[i]);
            }
            if (last >= y1) {
                return true;
            }
        }
        return false;
    }

    /** Accept a candidate range of the columns x1 to x2 and the rows y1 to y2. */
    add(x1: number, x2: number, y1: number, y2: number) {
        for (let x = x1; x <= x2; x++) {
            const column = this.columns.get(x)!;
            for (let i = CoveredColumns.countUpTo(column.starts, y1); i < column.ends.length; i += i & -i) {
                column.ends[i] = Math.max(column.ends[i], y2);
            }
        }
    }
}

/**
 * The merged ranges that can be ranges of the sheet, and the names of the others, as the other modes check them. The
 * ranges of keep that merges has with the same spans stay first. Then each entry of merges, in the order of the keys,
 * stays when mergeProblem accepts it and it does not cross a range that stays.
 */
function checkMerges(merges: Record<string, readonly unknown[]> | undefined, columns: number, rows: number, keep: Record<string, [number, number]> = {}) {
    const refused: string[] = [];
    const result: Record<string, [number, number]> = {};
    // The entries that mergeProblem accepts, with their first column, last column, first row and last row
    const candidates = new Map<string, [number, number, number, number]>();
    const entries = Object.entries(merges || {});
    for (const [name, span] of entries) {
        if (!mergeProblem(name, span, columns, rows)) {
            const [colspan, rowspan] = span as number[];
            const [x1, y1] = getIdFromColumnName(name);
            candidates.set(name, [x1, x1 + colspan - 1, y1, y1 + rowspan - 1]);
        }
    }
    const covered = new CoveredColumns(Array.from(candidates.values(), ([x1, x2, y1]) => [x1, x2, y1] as const));
    function accept(name: string) {
        const [x1, x2, y1, y2] = candidates.get(name)!;
        covered.add(x1, x2, y1, y2);
        result[name] = [x2 - x1 + 1, y2 - y1 + 1];
    }
    const unchanged = new Set<string>();
    for (const [name, span] of entries) {
        const kept = keep[name];
        if (kept && candidates.has(name) && span[0] === kept[0] && span[1] === kept[1]) {
            unchanged.add(name);
            accept(name);
        }
    }
    for (const [name] of entries) {
        if (unchanged.has(name)) {
            continue;
        }
        const range = candidates.get(name);
        if (!range || covered.crosses(...range)) {
            refused.push(name);
            continue;
        }
        accept(name);
    }
    // The ranges keep their order of the keys
    const ordered: Record<string, [number, number]> = {};
    for (const [name] of entries) {
        if (result[name]) {
            ordered[name] = result[name];
        }
    }
    return { merges: ordered, refused };
}

/**
 * The rows of a worksheet that have a record: all rows, or the frozen rows and the window of the row source mode.
 */
export function recordRows(worksheet: Worksheet) {
    return worksheet.rowSource ? worksheet.rowSource.attachedRows() : worksheet.rows.map((row, y) => y);
}

/**
 * The row source mode of a worksheet: the application gives the row count, the row heights and the hidden rows as runs,
 * and the values of each row when the window needs it. The arrays rows, records and options.data have one item for
 * each row, but only the frozen rows and the rows of the window have an item. The table body holds the frozen rows, a
 * spacer row, the rows of the window and a second spacer row. The spacer rows give the scroll area the height of all
 * rows. Above maxScrollHeight, a scroll position maps to a row in proportion.
 *
 * Data.set makes it from the worksheet option rowSource, as worksheet.rowSource.
 */
export class RowSource {
    readonly worksheet: Worksheet;

    count: number;

    layout: RowLayout;

    /** The merged ranges of the sheet. The window merges the part of a range that it holds. */
    merges: Record<string, [number, number]>;

    /** The merged ranges sorted by their first row, and the largest last row up to each range, for a fast lookup. */
    private mergeList: { name: string; x1: number; y1: number; x2: number; y2: number }[] = [];
    private mergeEnds: number[] = [];

    /** The first and the last row of the window after the frozen rows. The window is empty when to is below from. */
    from = 0;
    to = -1;

    private readonly source: jspreadsheet.RowSource;
    private readonly buffer: number;
    private readonly maxScrollHeight: number;
    private readonly top: HTMLTableRowElement;
    private readonly bottom: HTMLTableRowElement;
    private theadHeight = 0;
    private expectedTop: number | null = null;
    private wantedTop: number | null = null;
    private fillTimer: ReturnType<typeof setTimeout> | null = null;
    private observer: ResizeObserver | null = null;
    /** The height hook gave a row another height, so the row at the top of the view must stay there. */
    private changed = false;
    private measured = false;
    private focusLost = false;
    /** The focus that the table got from a cell that left the window is not lost any more when it leaves the table. */
    private readonly onFocusOut = (e: FocusEvent) => {
        if (e.target === this.worksheet.table) {
            this.focusLost = false;
        }
    };

    constructor(worksheet: Worksheet, source: jspreadsheet.RowSource) {
        this.worksheet = worksheet;
        this.source = source;
        this.count = Math.max(1, Math.floor(source.count) || 0);
        this.layout = new RowLayout(this.count, source.defaultHeight || Number(worksheet.options.defaultRowHeight) || 20, source.runs);
        this.buffer = source.buffer! > 0 ? source.buffer! : defaultBuffer;
        this.maxScrollHeight = source.maxScrollHeight! > 0 ? Math.min(source.maxScrollHeight!, defaultMaxScrollHeight) : defaultMaxScrollHeight;
        this.merges = checkMerges(worksheet.options.mergeCells as Record<string, unknown[]> | undefined, worksheet.headers.length, this.count).merges;
        this.indexMerges();
        this.top = this.spacer();
        this.bottom = this.spacer();
    }

    /** Fill the table body with the frozen rows and the spacers, and render the window. */
    start() {
        const worksheet = this.worksheet;
        worksheet.options.mergeCells = {};
        worksheet.rows = new Array(this.count);
        worksheet.records = new Array(this.count);
        worksheet.options.data = new Array(this.count);
        worksheet.tbody.innerHTML = '';
        for (let y = 0; y < this.frozenCount(); y++) {
            this.attachRow(y);
            worksheet.tbody.appendChild(worksheet.rows[y].element);
        }
        worksheet.tbody.appendChild(this.top);
        worksheet.tbody.appendChild(this.bottom);
        worksheet.table.addEventListener('focusout', this.onFocusOut);
        this.applyMerges();
        this.render(true);
        // The window renders again when the view changes its size, for example when the tab of the worksheet opens
        if (typeof ResizeObserver === 'function') {
            this.observer = new ResizeObserver(() => {
                if (worksheet.rowSource === this && worksheet.content.isConnected) {
                    this.render();
                }
            });
            this.observer.observe(worksheet.content);
        }
    }

    /** Stop the size observer and the pending fill, for example before the worksheet goes. */
    stop() {
        this.worksheet.table.removeEventListener('focusout', this.onFocusOut);
        this.observer?.disconnect();
        this.observer = null;
        if (this.fillTimer) {
            clearTimeout(this.fillTimer);
            this.fillTimer = null;
        }
    }

    /** The rows with a record, in order: the frozen rows and the rows of the window. */
    attachedRows() {
        const rows: number[] = [];
        const frozen = this.frozenCount();
        for (let y = 0; y < frozen; y++) {
            if (this.worksheet.rows[y]) {
                rows.push(y);
            }
        }
        for (let y = Math.max(frozen, this.from); y <= this.to; y++) {
            if (this.worksheet.rows[y]) {
                rows.push(y);
            }
        }
        return rows;
    }

    private frozenCount() {
        return Math.min(this.worksheet.options.freezeRows || 0, this.count);
    }

    private spacer() {
        const tr = document.createElement('tr');
        tr.className = 'jss_spacer';
        tr.setAttribute('aria-hidden', 'true');
        tr.setAttribute('role', 'presentation');
        const td = document.createElement('td');
        td.setAttribute('colspan', this.worksheet.headers.length + 1);
        td.style.cssText = 'padding:0;border:0;height:0;background:transparent';
        tr.appendChild(td);
        tr.style.display = 'none';
        return tr;
    }

    private setSpacer(tr: HTMLTableRowElement, height: number) {
        height = Math.max(0, Math.round(height));
        tr.style.display = height > 0 ? '' : 'none';
        tr.style.height = `${height}px`;
        (tr.firstChild as HTMLElement).style.height = `${height}px`;
    }

    /** Make the record of row y from the values of the application. */
    private attachRow(y: number) {
        const worksheet = this.worksheet;
        const columns = worksheet.headers.length;
        const given = this.source.row(y);
        const data: jspreadsheet.CellValue[] = new Array(columns);
        for (let x = 0; x < columns; x++) {
            const value = given ? given[x] : undefined;
            data[x] = value === undefined || value === null ? '' : value;
        }
        worksheet.options.data[y] = data;
        const row = worksheet.rowManager.create(y, data);
        // A cell of a hidden column is hidden
        for (let x = 0; x < columns; x++) {
            if (worksheet.headers[x].style.display === 'none') {
                worksheet.records[y][x].element.style.display = 'none';
            }
        }
        // The application gives the height of the row that shows, for example from its content
        if (typeof this.source.height === 'function') {
            const height = this.source.height(y, data);
            if (height > 0 && height !== this.layout.shownHeight(y)) {
                this.layout.set(y, y, { height });
                this.changed = true;
            }
        }
        row.element.style.height = `${this.layout.shownHeight(y)}px`;
        if (this.layout.isHidden(y)) {
            row.element.style.display = 'none';
        }
        return row;
    }

    /** Remove the record of row y. An open editor in the row detaches first, as oneditordetach describes. */
    private detachRow(y: number) {
        const worksheet = this.worksheet;
        const row = worksheet.rows[y];
        if (!row) {
            return;
        }
        if (worksheet.editor.isOpen() && row.element.contains(worksheet.editor.cell)) {
            worksheet.editor.detach();
        }
        row.element.remove();
        delete worksheet.rows[y];
        delete worksheet.records[y];
        delete worksheet.options.data[y];
    }

    /** The scroll position of the view in grid pixels. */
    private scrollTop() {
        return toGrid(this.worksheet, this.worksheet.content.scrollTop);
    }

    private setScrollTop(top: number) {
        this.worksheet.content.scrollTop = toScreen(this.worksheet, top);
    }

    private geometry(): Geometry {
        const worksheet = this.worksheet;
        const layout = this.layout;
        const frozen = this.frozenCount();
        const frozenHeight = layout.offsetOf(frozen);
        if (!this.theadHeight) {
            this.theadHeight = worksheet.thead.offsetHeight || 0;
        }
        // A view without a height limit would hold all rows, so the view is at most the height of the screen
        const screenHeight = toGrid(worksheet, Math.max(1000, window.innerHeight || 0));
        const view = Math.max(1, Math.min(toGrid(worksheet, worksheet.content.clientHeight), screenHeight) - this.theadHeight - frozenHeight);
        const total = layout.total;
        // Above 100 %, the limit in grid pixels is lower, so the scroll area stays below the limit of the browser
        const maxHeight = this.maxScrollHeight / Math.max(zoomOf(worksheet), 1);
        const scaled = total > maxHeight;
        const flow = scaled ? maxHeight - frozenHeight : total - frozenHeight;
        const maxScroll = Math.max(1, flow - view);
        const maxTop = Math.max(0, total - frozenHeight - view);
        return { frozen, frozenHeight, view, total, scaled, flow, maxScroll, maxTop };
    }

    /** The place in the rows of the top of the view after the frozen rows. */
    private viewTop(g: Geometry) {
        const scrollTop = this.scrollTop();
        if (this.wantedTop !== null && this.expectedTop !== null && Math.abs(scrollTop - this.expectedTop) < 2) {
            return this.wantedTop;
        }
        this.wantedTop = null;
        this.expectedTop = null;
        if (!g.scaled) {
            return g.frozenHeight + scrollTop;
        }
        return g.frozenHeight + (Math.min(scrollTop, g.maxScroll) / g.maxScroll) * g.maxTop;
    }

    /** The scroll position that shows the place top of the rows at the top of the view. */
    private scrollFor(g: Geometry, top: number) {
        const offset = Math.max(0, Math.min(top - g.frozenHeight, g.maxTop));
        return g.scaled ? (offset / g.maxTop) * g.maxScroll : offset;
    }

    /**
     * Render the window of rows for the scroll position. With force, make a new window also when the view is inside the
     * window. With scrolling (a scroll event), a view without a row of the window is a far scroll: the window gets few
     * rows now, and all rows of the buffer when no far scroll comes for a short time. onrowwindow occurs when the window
     * changes.
     */
    render(force?: boolean, scrolling?: boolean) {
        const worksheet = this.worksheet;
        if (!worksheet.content) {
            return;
        }
        const layout = this.layout;
        const g = this.geometry();
        const scrollBefore = worksheet.content.scrollTop;
        const top = this.viewTop(g);
        const first = Math.max(g.frozen, layout.rowAt(top));
        const last = Math.max(first, Math.min(this.count - 1, layout.rowAt(top + g.view)));
        // The place of the view in its first row, to keep it after the application changes row heights
        const within = top - layout.offsetOf(first);
        const margin = this.buffer >> 2;
        let from = this.from;
        let to = this.to;
        const inside = to >= from && from <= Math.max(g.frozen, first - margin) && to >= Math.min(this.count - 1, last + margin);
        const far = !!scrolling && !force && !inside && (to < from || first > to || last < from);
        if (force || !inside) {
            const buffer = far ? Math.min(this.buffer, farScrollBuffer) : this.buffer;
            from = Math.max(g.frozen, first - buffer);
            to = Math.min(this.count - 1, last + buffer);
            from = this.mergeReach(from, to);
            if (far) {
                this.scheduleFill();
            } else if (this.fillTimer) {
                clearTimeout(this.fillTimer);
                this.fillTimer = null;
            }
        }
        // The spacer above the window puts the row at top at the top of the view. The browser can scroll a few pixels
        // past maxScroll, because the scroll height holds the border of the view, so those pixels do not count.
        const scrollTop = g.scaled ? Math.min(this.scrollTop(), g.maxScroll) : this.scrollTop();
        let above = scrollTop - top + layout.offsetOf(from);
        if (above < 0) {
            // Near the top of a scaled scroll area, the rows above the view do not fit above it
            const limit = top - scrollTop;
            let row = layout.rowAt(limit);
            if (layout.offsetOf(row) < limit) {
                row++;
            }
            from = Math.min(Math.max(g.frozen, row), first);
            above = Math.max(0, scrollTop - top + layout.offsetOf(from));
        }
        let below = g.flow - above - (layout.offsetOf(to + 1) - layout.offsetOf(from));
        if (below < 0) {
            const end = layout.offsetOf(from) + g.flow - above;
            to = Math.max(last, Math.min(to, layout.rowAt(end) - 1));
            below = Math.max(0, g.flow - above - (layout.offsetOf(to + 1) - layout.offsetOf(from)));
        }
        const changed = from !== this.from || to !== this.to;
        if (changed || ((force || this.activeRowMissing(from, to)) && this.missingRows(from, to).length > 0)) {
            this.moveWindow(from, to);
        }
        this.setSpacer(this.top, above);
        this.setSpacer(this.bottom, below);
        if (this.changed || (typeof this.source.height !== 'function' && !this.measured && this.measureMinHeight())) {
            // The rows have other heights than the layout had. The row at the top of the view stays there.
            this.changed = false;
            this.keepRowAtTop(first, within, !far);
            if (changed && this.from === from && this.to === to) {
                worksheet.dispatch('onrowwindow', worksheet, from, to, first, last);
            }
            return;
        }
        if ((changed || force) && !far) {
            this.alignWindow(first, above, g);
        }
        // While the window changes its rows, the table can be shorter for a moment, and the browser lowers the scroll
        // position. The spacers are for the scroll position before the change, so it comes back.
        if (changed && worksheet.content.scrollTop !== scrollBefore) {
            worksheet.content.scrollTop = scrollBefore;
        }
        if (changed) {
            worksheet.dispatch('onrowwindow', worksheet, from, to, first, last);
        }
    }

    private scheduleFill() {
        if (this.fillTimer) {
            clearTimeout(this.fillTimer);
        }
        this.fillTimer = setTimeout(() => {
            this.fillTimer = null;
            if (this.worksheet.rowSource === this && this.worksheet.content?.isConnected) {
                this.render(true);
            }
        }, fillDelay);
    }

    /**
     * Measure the height of a row with the default height once the table shows. When the cells make it higher than the
     * layout, for example with their padding, the layout gets that height as its smallest height.
     * @returns True when the layout changed.
     */
    private measureMinHeight() {
        const worksheet = this.worksheet;
        if (!worksheet.content.isConnected || !worksheet.content.clientHeight) {
            return false;
        }
        for (let y = this.from; y <= this.to; y++) {
            const row = worksheet.rows[y];
            if (!row || this.layout.runOf(y) || !row.element.offsetParent) {
                continue;
            }
            this.measured = true;
            const height = toGrid(worksheet, row.element.getBoundingClientRect().height);
            if (height > this.layout.defaultHeight + 0.5) {
                this.layout.setMinHeight(height);
                return true;
            }
            return false;
        }
        return false;
    }

    /** Scroll so that row y is at the top of the view, within pixels below its top, and render the window again. */
    private keepRowAtTop(y: number, within: number, force: boolean) {
        const g = this.geometry();
        const offset = this.layout.offsetOf(y) + Math.max(0, Math.min(within || 0, this.layout.heightOf(y) - 1));
        const wanted = Math.max(g.frozenHeight, Math.min(offset, g.frozenHeight + g.maxTop));
        this.setScrollTop(this.scrollFor(g, wanted));
        this.expectedTop = this.scrollTop();
        this.wantedTop = wanted;
        this.render(force);
    }

    /** The row at the top of the view after the frozen rows, and the place of the view in it. */
    private viewRow() {
        const g = this.geometry();
        const top = this.viewTop(g);
        const first = Math.max(g.frozen, this.layout.rowAt(top));
        return { first, within: top - this.layout.offsetOf(first) };
    }

    /** The top of a table row. A row that does not show has no place, so it has the bottom of the row before it. */
    private topOf(tr: HTMLTableRowElement) {
        if (tr.style.display !== 'none') {
            return tr.offsetTop;
        }
        for (let row = tr.previousElementSibling as HTMLElement | null; row; row = row.previousElementSibling as HTMLElement | null) {
            if (row.offsetParent) {
                return row.offsetTop + row.offsetHeight;
            }
        }
        return this.worksheet.tbody.offsetTop;
    }

    /**
     * Correct the spacers when the rows of the window in the browser are higher or lower than in the layout, for example
     * a row with wrapped text. The row first stays at the top of the view, and the scroll area keeps its height.
     */
    private alignWindow(first: number, above: number, g: Geometry) {
        const worksheet = this.worksheet;
        const row = worksheet.rows[first];
        if (!row?.element.offsetParent || !worksheet.content.clientHeight) {
            return;
        }
        const start = this.topOf(this.top) + (this.top.style.display === 'none' ? 0 : this.top.offsetHeight);
        const end = this.topOf(this.bottom);
        const real = row.element.offsetTop - start;
        const expected = this.layout.offsetOf(first) - this.layout.offsetOf(this.from);
        const height = end - start;
        if (Math.abs(real - expected) < 1 && Math.abs(height - (this.layout.offsetOf(this.to + 1) - this.layout.offsetOf(this.from))) < 1) {
            return;
        }
        const corrected = Math.max(0, above - (real - expected));
        this.setSpacer(this.top, corrected);
        this.setSpacer(this.bottom, Math.max(0, g.flow - corrected - height));
    }

    /** Sort the merged ranges by their first row for mergesIn. */
    private indexMerges() {
        this.mergeList = Object.entries(this.merges)
            .map(([name, [colspan, rowspan]]) => {
                const [x1, y1] = getIdFromColumnName(name);
                return { name, x1, y1, x2: x1 + colspan - 1, y2: y1 + rowspan - 1 };
            })
            .sort((a, b) => a.y1 - b.y1);
        let end = -1;
        this.mergeEnds = this.mergeList.map((range) => (end = Math.max(end, range.y2)));
    }

    /** The merged ranges that cross the rows `from` to `to`, without a loop over all ranges. */
    mergesIn(from: number, to: number) {
        const list = this.mergeList;
        // The last range whose first row is at most `to`
        let low = 0;
        let high = list.length - 1;
        let last = -1;
        while (low <= high) {
            const middle = (low + high) >> 1;
            if (list[middle].y1 <= to) {
                last = middle;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        const result = [];
        for (let i = last; i >= 0 && this.mergeEnds[i] >= from; i--) {
            if (list[i].y2 >= from) {
                result.push(list[i]);
            }
        }
        return result;
    }

    /** Move the start of the window up to the first row of each merged range that covers its first row. */
    private mergeReach(from: number, to: number) {
        let result = from;
        for (const range of this.mergesIn(from, to)) {
            if (range.y1 < result && from - range.y1 <= maxMergeReach) {
                result = range.y1;
            }
        }
        return Math.max(this.frozenCount(), result);
    }

    /**
     * The rows from `from` to `to` without a record that the window must attach. A hidden row gets no record, so a long
     * run of hidden rows costs nothing. A hidden row of a merged range of the window gets one, so the range keeps its
     * rows, and so does a hidden row of the selection, so that the keys can start from it.
     */
    private missingRows(from: number, to: number) {
        const worksheet = this.worksheet;
        const merged = this.mergedRows(from, to);
        for (const y of this.activeRows()) {
            if (y >= from && y <= to) {
                merged.push([y, y]);
            }
        }
        const result: number[] = [];
        for (let y = from; y <= to; y++) {
            if (worksheet.rows[y]) {
                continue;
            }
            const run = this.layout.isHidden(y) ? this.layout.runOf(y) : null;
            if (run) {
                const end = Math.min(to, run.last);
                for (const [a, b] of merged) {
                    for (let row = Math.max(y, a); row <= Math.min(end, b); row++) {
                        if (!worksheet.rows[row]) {
                            result.push(row);
                        }
                    }
                }
                y = end;
                continue;
            }
            result.push(y);
        }
        return result.sort((a, b) => a - b).filter((y, i, list) => i === 0 || list[i - 1] !== y);
    }

    /** The rows of the two corners of the selection. */
    private activeRows() {
        const range = this.worksheet.selection.range;
        return range ? [range[1], range[3]].filter((y) => y >= 0) : [];
    }

    private activeRowMissing(from: number, to: number) {
        return this.activeRows().some((y) => y >= from && y <= to && !this.worksheet.rows[y]);
    }

    /** The rows of the merged ranges that cross the rows `from` to `to`, as [first, last] pairs. */
    private mergedRows(from: number, to: number) {
        const result: [number, number][] = [];
        for (const range of this.mergesIn(from, to)) {
            if (range.y2 > range.y1) {
                result.push([Math.max(from, range.y1), Math.min(to, range.y2)]);
            }
        }
        return result;
    }

    /** Make the rows `from` to `to` the window: remove the other rows, and add the missing rows in order. */
    private moveWindow(from: number, to: number) {
        const worksheet = this.worksheet;
        this.clearMerges();
        const removed = new Set<number>();
        for (let y = this.from; y <= this.to; y++) {
            if ((y < from || y > to) && worksheet.rows[y]) {
                removed.add(y);
                // The focus in the row, also in its editor, goes to the table. The application takes the edit before the
                // editor closes, so the blur of the editor does not save it.
                const element = worksheet.rows[y].element;
                const focused = !!document.activeElement && element.contains(document.activeElement);
                if (worksheet.editor.isOpen() && element.contains(worksheet.editor.cell)) {
                    worksheet.editor.detach();
                }
                if (focused) {
                    this.focusTable();
                }
                this.detachRow(y);
            }
        }
        if (removed.size && worksheet.selection.highlighted.length) {
            worksheet.selection.highlighted = worksheet.selection.highlighted.filter((record) => record.element.isConnected);
        }
        this.dropChain(removed);
        const added = this.missingRows(from, to);
        // Each new row goes before the next row of the window with a record, else before the bottom spacer
        let next: HTMLTableRowElement = this.bottom;
        let bound = to + 1;
        let fragment = document.createDocumentFragment();
        for (let i = added.length - 1; i >= 0; i--) {
            const y = added[i];
            const after = this.nextAttached(y + 1, bound - 1);
            if (after !== null) {
                next.before(fragment);
                fragment = document.createDocumentFragment();
                next = worksheet.rows[after].element;
            }
            bound = y;
            this.attachRow(y);
            fragment.insertBefore(worksheet.rows[y].element, fragment.firstChild);
        }
        next.before(fragment);
        this.from = from;
        this.to = to;
        this.applyMerges();
        if (added.length) {
            this.decorate(ranges(added));
        }
        this.updateTabStop();
    }

    /** Remove the formula cells of removed rows from the formula chain. Their formulas add them again when they show. */
    private dropChain(rows: Set<number>) {
        if (!rows.size) {
            return;
        }
        const chain = this.worksheet.formulas.chain;
        for (const name of Object.keys(chain)) {
            chain[name] = chain[name].filter((dependent) => !rows.has(getIdFromColumnName(dependent)[1]));
            if (!chain[name].length) {
                delete chain[name];
            }
        }
    }

    /** The table is the tab stop while the row of the active cell is outside the window, so the keys reach the grid. */
    private updateTabStop() {
        const worksheet = this.worksheet;
        let cell = worksheet.selection.activeCell;
        // A cell in a hidden row, or a cell that a merged range hides, cannot get the focus, so a cell that shows gets the tab stop
        if (
            cell?.isConnected &&
            worksheet.tbody.contains(cell) &&
            (cell.style.display === 'none' || (cell.parentNode as HTMLElement).style.display === 'none')
        ) {
            // A window change does not drop the focus to the body, so the focus does not move here
            worksheet.selection.updateTabStop(false);
            cell = worksheet.selection.activeCell;
            // The focus that the table got from a cell that left the window goes to the new tab stop
            if (this.focusLost && document.activeElement === worksheet.table && cell && shows(cell) && worksheet.tbody.contains(cell)) {
                this.focusLost = false;
                cell.focus({ preventScroll: true });
            }
        }
        worksheet.table.tabIndex = cell?.isConnected && worksheet.table.contains(cell) ? -1 : 0;
    }

    /** The first row from `from` to `to` with a record, or null. The runs of hidden rows are skipped. */
    private nextAttached(from: number, to: number) {
        const worksheet = this.worksheet;
        for (let y = from; y <= to; y++) {
            if (worksheet.rows[y]) {
                return y;
            }
            const run = this.layout.isHidden(y) ? this.layout.runOf(y) : null;
            if (run) {
                // A hidden row of the run can have a record: it was in the window before it became hidden
                for (let row = y + 1; row <= Math.min(to, run.last); row++) {
                    if (worksheet.rows[row]) {
                        return row;
                    }
                }
                y = run.last;
            }
        }
        return null;
    }

    /** Give new rows the selection and the frozen columns, give the focus back, and tell the application. */
    private decorate(added: [number, number][]) {
        const worksheet = this.worksheet;
        for (const [a, b] of added) {
            for (let y = a; y <= b; y++) {
                this.selectRow(y);
            }
        }
        // The frozen columns of the new rows get their place, from the scroll position
        if (worksheet.options.freezeColumns) {
            worksheet.columns.updateFreezePosition();
        }
        this.restoreFocus();
        worksheet.dispatch('onrowsattached', worksheet, added);
    }

    /** Remove the merged ranges of the window from the table. */
    private clearMerges() {
        for (const name of Object.keys(this.worksheet.options.mergeCells || {})) {
            this.worksheet.merges.unmergeWindow(name);
        }
    }

    /** Merge the part of each merged range whose first row has a record. A range ends at the last row with a record. */
    private applyMerges() {
        const worksheet = this.worksheet;
        const frozen = this.frozenCount();
        // The ranges whose first row can have a record: in the frozen rows or in the window
        const ranges = [...(frozen ? this.mergesIn(0, frozen - 1) : []), ...(this.to >= this.from ? this.mergesIn(this.from, this.to) : [])];
        const done = new Set<string>();
        for (const { name, x1: x, y1: y } of ranges) {
            if (done.has(name) || !worksheet.rows[y]) {
                continue;
            }
            done.add(name);
            const span = this.merges[name];
            let end = Math.min(y + span[1] - 1, y < frozen && this.from > frozen ? frozen - 1 : this.to);
            if (y < frozen && this.to < this.from) {
                end = Math.min(end, frozen - 1);
            }
            // A hidden row of the window without a record ends the range
            for (let row = y + 1; row <= end; row++) {
                if (!worksheet.records[row]) {
                    end = row - 1;
                    break;
                }
            }
            const rowspan = Math.max(end, y) - y + 1;
            const colspan = Math.min(span[0], worksheet.headers.length - x);
            if ((colspan > 1 || rowspan > 1) && worksheet.merges.mergeWindow(name, colspan, rowspan)) {
                // The first cell shows when one of the columns of the range shows
                const element = worksheet.records[y][x].element;
                if (element.style.display === 'none') {
                    for (let i = x; i < x + colspan; i++) {
                        if (worksheet.headers[i]?.style.display !== 'none') {
                            element.style.display = '';
                            break;
                        }
                    }
                }
            }
        }
    }

    /** Give row y the classes of the selection, and the tab stop when it has the selected cell. */
    selectRow(y: number) {
        const worksheet = this.worksheet;
        const range = worksheet.selection.range;
        const bounds = worksheet.selection.bounds;
        if (!range || !bounds) {
            return;
        }
        const [left, topRow, right, bottomRow] = bounds;
        const records = worksheet.records[y];
        if (y < topRow || y > bottomRow || !records) {
            return;
        }
        const hidden = worksheet.rows[y].element.style.display === 'none';
        for (let x = left; x <= right; x++) {
            const record = records[x];
            if (!record || hidden || record.element.style.display === 'none') {
                continue;
            }
            record.element.classList.add('highlight');
            record.element.setAttribute('aria-selected', 'true');
            worksheet.selection.highlighted.push(record);
            if (y === topRow) {
                record.element.classList.add('highlight-top');
            }
            if (y === bottomRow) {
                record.element.classList.add('highlight-bottom');
            }
        }
        if (!hidden) {
            records[left]?.element.classList.add('highlight-left');
            records[right]?.element.classList.add('highlight-right');
            worksheet.rows[y].element.classList.add('selected');
        }
        if (y === range[1] && records[range[0]]) {
            const td = records[range[0]].element;
            td.classList.add('highlight-selected');
            worksheet.selection.setActiveCell(td);
        }
    }

    /** The focus in an element that the window removes goes to the table, so that the keys keep working. */
    private keepFocus(element: HTMLElement) {
        const active = document.activeElement;
        if (active && element.contains(active)) {
            this.focusTable();
        }
    }

    private focusTable() {
        const worksheet = this.worksheet;
        if (!worksheet.table.hasAttribute('tabindex')) {
            worksheet.table.tabIndex = -1;
        }
        worksheet.table.focus({ preventScroll: true });
        this.focusLost = true;
    }

    /** Give the focus back to the active cell when its row is in the window again. */
    private restoreFocus() {
        const worksheet = this.worksheet;
        const range = worksheet.selection.range;
        if (!this.focusLost || document.activeElement !== worksheet.table || !range) {
            return;
        }
        const record = worksheet.records[range[1]]?.[range[0]];
        // A cell in a hidden row cannot get the focus: updateTabStop gives it to the new tab stop
        if (record && shows(record.element)) {
            record.element.focus({ preventScroll: true });
            // A cell that is not the tab stop yet cannot get the focus, for example a frozen row before its selection
            this.focusLost = document.activeElement !== record.element;
        }
    }

    /**
     * Scroll so that row y shows, and render its window. The row goes to the top of the view, or to the bottom when
     * toBottom is true. A row that shows does not move.
     */
    scrollToRow(y: number, toBottom?: boolean) {
        y = Math.max(0, Math.min(this.count - 1, Math.floor(Number(y)) || 0));
        const g = this.geometry();
        if (y < g.frozen) {
            this.render();
            return;
        }
        const top = this.viewTop(g);
        const rowTop = this.layout.offsetOf(y);
        const rowBottom = this.layout.offsetOf(y + 1);
        let wanted = top;
        if (rowTop < top) {
            wanted = rowTop;
        } else if (rowBottom > top + g.view) {
            wanted = toBottom === false ? rowTop : rowBottom - g.view;
        }
        if (wanted !== top) {
            wanted = Math.max(g.frozenHeight, Math.min(wanted, g.frozenHeight + g.maxTop));
            this.setScrollTop(this.scrollFor(g, wanted));
            this.expectedTop = this.scrollTop();
            this.wantedTop = wanted;
        }
        this.render();
    }

    /** The top of row y in the layout, in grid pixels. */
    rowOffset(y: number) {
        return this.layout.offsetOf(y);
    }

    /**
     * Make the records of the rows `from` to `to` again from the values of the application, for example after it loaded
     * them. Rows outside the window do not change.
     */
    refreshRows(from: number, to: number) {
        const worksheet = this.worksheet;
        const frozen = this.frozenCount();
        from = Math.max(0, Math.floor(from) || 0);
        to = Math.min(this.count - 1, Math.floor(to) || 0);
        const rows: number[] = [];
        for (let y = from; y <= to; y++) {
            if (worksheet.rows[y] && (y < frozen || (y >= this.from && y <= this.to))) {
                rows.push(y);
            }
        }
        if (!rows.length) {
            return;
        }
        const view = this.viewRow();
        this.clearMerges();
        for (const y of rows) {
            const old = worksheet.rows[y].element;
            const focused = !!document.activeElement && old.contains(document.activeElement);
            if (worksheet.editor.isOpen() && old.contains(worksheet.editor.cell)) {
                worksheet.editor.detach();
            }
            if (focused) {
                this.focusTable();
            }
            delete worksheet.records[y];
            this.attachRow(y);
            old.replaceWith(worksheet.rows[y].element);
        }
        worksheet.selection.highlighted = worksheet.selection.highlighted.filter((record) => record.element.isConnected);
        this.applyMerges();
        this.decorate(ranges(rows));
        if (this.changed) {
            this.changed = false;
            this.keepRowAtTop(view.first, view.within, false);
        }
    }

    /**
     * Ask the application again for the height of each row of rows that has a record, for example after a change of its
     * cells. The row at the top of the view stays there.
     */
    fitRows(rows: number[]) {
        const worksheet = this.worksheet;
        if (typeof this.source.height !== 'function') {
            return;
        }
        const view = this.viewRow();
        let changed = false;
        for (const value of rows) {
            const y = Math.floor(value);
            const row = worksheet.rows[y];
            if (!row) {
                continue;
            }
            const height = this.source.height(y, worksheet.options.data[y]);
            if (height > 0 && height !== this.layout.shownHeight(y)) {
                this.layout.set(y, y, { height });
                row.element.style.height = `${height}px`;
                changed = true;
            }
        }
        if (changed) {
            this.keepRowAtTop(view.first, view.within, false);
        }
    }

    /**
     * Replace the row count, the runs and the merged ranges, for example after the application changed its rows. All
     * records are made again. The scroll position and the selection stay. When a merged range hides the first selected
     * cell, a selection inside the range selects its first cell, and a larger selection grows to hold the range.
     */
    reset(options?: { count?: number; runs?: jspreadsheet.RowRun[]; merges?: Record<string, [number, number]> }) {
        const worksheet = this.worksheet;
        const view = this.viewRow();
        this.clearMerges();
        if (document.activeElement && worksheet.tbody.contains(document.activeElement)) {
            this.focusTable();
        }
        if (worksheet.editor.isOpen()) {
            worksheet.editor.detach();
        }
        for (const y of this.attachedRows()) {
            this.keepFocus(worksheet.rows[y].element);
            this.detachRow(y);
        }
        worksheet.selection.highlighted = [];
        const count = Math.max(1, Math.floor(options?.count ?? this.count) || 1);
        this.count = count;
        this.layout = new RowLayout(count, this.layout.givenHeight, options?.runs ?? this.layout.runs, this.layout.minHeight);
        // The merged ranges must fit the new row count also without new ranges
        this.merges = checkMerges(options?.merges ?? this.merges, worksheet.headers.length, count, this.merges).merges;
        this.indexMerges();
        worksheet.rows = new Array(count);
        worksheet.records = new Array(count);
        worksheet.options.data = new Array(count);
        this.from = 0;
        this.to = -1;
        const frozen = this.frozenCount();
        const frozenRows = document.createDocumentFragment();
        for (let y = 0; y < frozen; y++) {
            this.attachRow(y);
            frozenRows.appendChild(worksheet.rows[y].element);
        }
        this.top.before(frozenRows);
        const range = worksheet.selection.range;
        if (range) {
            range[1] = Math.min(range[1], count - 1);
            range[3] = Math.min(range[3], count - 1);
            const bounds = worksheet.selection.bounds;
            if (bounds) {
                bounds[1] = Math.min(bounds[1], count - 1);
                bounds[3] = Math.min(bounds[3], count - 1);
            }
        }
        worksheet.accessibility.updateGridSize();
        this.setSpacer(this.top, 0);
        this.setSpacer(this.bottom, this.geometry().flow);
        this.keepRowAtTop(Math.min(view.first, count - 1), view.within, true);
        if (frozen) {
            this.applyMerges();
            this.decorate([[0, frozen - 1]]);
            this.updateTabStop();
        }
        const selected = worksheet.selection.range?.map(Number);
        const merged = selected && this.mergeAt(selected[0], selected[1]);
        if (merged && (merged[0] !== selected[0] || merged[1] !== selected[1])) {
            const inside = selected[2] >= merged[0] && selected[2] <= merged[2] && selected[3] >= merged[1] && selected[3] <= merged[3];
            // A selection inside the range goes to its first cell. A larger selection stays as given, and grows to hold
            // the range.
            if (inside) {
                worksheet.selection.updateFromCoords(merged[0], merged[1], merged[0], merged[1]);
            } else {
                worksheet.selection.updateFromCoords(selected[0], selected[1], selected[2], selected[3]);
            }
        }
        // The focus of a cell that is in a frozen row now goes back to it
        this.restoreFocus();
    }

    /** Set the height (pixels) or the hidden state of rows in the layout and in their records. */
    setLayout(rows: number[], values: { height?: number | null; hidden?: boolean }) {
        const worksheet = this.worksheet;
        for (const value of rows) {
            const y = Math.floor(value);
            this.layout.set(y, y, values);
            const row = worksheet.rows[y];
            if (row) {
                if ('height' in values) {
                    row.element.style.height = `${this.layout.shownHeight(y)}px`;
                }
                if ('hidden' in values) {
                    row.element.style.display = values.hidden ? 'none' : '';
                }
            }
        }
        this.render(true);
    }

    /**
     * The element of the cell (x, y). When its row has no record, the window scrolls to it first. Null when the cell has
     * no record after the scroll, for example in a hidden row.
     */
    reveal(x: number, y: number) {
        const worksheet = this.worksheet;
        if (!worksheet.records[y]) {
            this.scrollToRow(y);
        }
        return worksheet.records[y]?.[x]?.element ?? null;
    }

    /** The merged range that holds cell (x, y), as [x1, y1, x2, y2], or null. */
    mergeAt(x: number, y: number): [number, number, number, number] | null {
        for (const range of this.mergesIn(y, y)) {
            if (x >= range.x1 && x <= range.x2) {
                return [range.x1, range.y1, range.x2, range.y2];
            }
        }
        return null;
    }

    /**
     * Replace the merged ranges of the sheet. A range that the checks of the other modes refuse is left out.
     * @returns The first cells of the ranges that are left out.
     */
    replaceMerges(merges: Record<string, [number, number]>) {
        const checked = checkMerges(merges, this.worksheet.headers.length, this.count, this.merges);
        this.clearMerges();
        this.merges = checked.merges;
        this.indexMerges();
        // A new range can hold hidden rows of the window without a record. The window adds them.
        if (this.to >= this.from && this.missingRows(this.from, this.to).length > 0) {
            this.moveWindow(this.from, this.to);
        } else {
            this.applyMerges();
        }
        return checked.refused;
    }
}
