import { translate } from '../utils/dictionary.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { MergeInfo, Worksheet } from '../types/internal';

/**
 * The merged cells of a worksheet. The ranges are in options.mergeCells: the colspan, the rowspan and the td elements
 * that the first cell covers.
 *
 * Each worksheet has one, as worksheet.merges. getMerge, setMerge, removeMerge and destroyMerge call it, and the moves
 * and inserts of rows and columns ask it which ranges a row or a column is in.
 */
export class Merges {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Is column merged
     */
    // The moves pass 1 or 0 for insertBefore
    isColMerged(x: number, insertBefore?: boolean | 0 | 1 | null) {
        const worksheet = this.worksheet;

        const cols = [];
        // Remove any merged cells
        if (worksheet.options.mergeCells) {
            for (const [key, [colspan]] of Object.entries(worksheet.options.mergeCells)) {
                const info = getIdFromColumnName(key);
                const x1 = info[0];
                const x2 = info[0] + (colspan > 1 ? colspan - 1 : 0);

                if (insertBefore == null) {
                    if (x1 <= x && x2 >= x) {
                        cols.push(key);
                    }
                } else {
                    if (insertBefore) {
                        if (x1 < x && x2 >= x) {
                            cols.push(key);
                        }
                    } else {
                        if (x1 <= x && x2 > x) {
                            cols.push(key);
                        }
                    }
                }
            }
        }

        return cols;
    }

    /**
     * Is rows merged
     */
    // The moves pass 1 or 0 for insertBefore
    isRowMerged(y: number, insertBefore?: boolean | 0 | 1 | null) {
        const worksheet = this.worksheet;

        const rows = [];
        // Remove any merged cells
        if (worksheet.options.mergeCells) {
            for (const [key, [, rowspan]] of Object.entries(worksheet.options.mergeCells)) {
                const info = getIdFromColumnName(key);
                const y1 = info[1];
                const y2 = info[1] + (rowspan > 1 ? rowspan - 1 : 0);

                if (insertBefore == null) {
                    if (y1 <= y && y2 >= y) {
                        rows.push(key);
                    }
                } else {
                    if (insertBefore) {
                        if (y1 < y && y2 >= y) {
                            rows.push(key);
                        }
                    } else {
                        if (y1 <= y && y2 > y) {
                            rows.push(key);
                        }
                    }
                }
            }
        }

        return rows;
    }

    /**
     * Update the first cell of a merged range for the columns that are hidden.
     * The record of each cell of a merged range points to the element of the first cell.
     * So the first cell is hidden only when each column of the range is hidden,
     * and its colspan is the number of columns of the range that are not hidden.
     */
    private updateFirstCell(cellName: string) {
        const worksheet = this.worksheet;

        const cell = getIdFromColumnName(cellName);
        const colspan = worksheet.options.mergeCells![cellName][0];

        let visible = 0;
        for (let x = cell[0]; x < cell[0] + colspan; x++) {
            if (worksheet.headers[x].style.display !== 'none') {
                visible++;
            }
        }

        const element = worksheet.records[cell[1]][cell[0]].element;
        if (visible === 0) {
            element.style.display = 'none';
        } else {
            element.style.display = '';
            if (colspan > 1) {
                element.setAttribute('colspan', visible);
            }
        }

        // A first cell in a column with the type hidden shows the range with the options of another column
        if (worksheet.options.columns?.[cell[0]]?.type === 'hidden') {
            worksheet.cells.render(cell[0], cell[1]);
        }
    }

    /**
     * Update the first cell of each merged range that covers one of the columns.
     */
    updateColumns(columns: number[]) {
        const worksheet = this.worksheet;

        if (!worksheet.options.mergeCells) {
            return;
        }

        for (const key of Object.keys(worksheet.options.mergeCells)) {
            const cell = getIdFromColumnName(key);
            const colspan = worksheet.options.mergeCells[key][0];

            for (const column of columns) {
                const x = parseInt(column);
                if (x >= cell[0] && x < cell[0] + colspan) {
                    this.updateFirstCell(key);
                    break;
                }
            }
        }
    }

    /**
     * Get a merged range, or all of them.
     * @param cellName - The first cell of the range. Without it, the ranges by first cell.
     */
    get(cellName?: string) {
        const worksheet = this.worksheet;

        let data: Record<string, [number, number]> | [number, number] | null = {};
        if (cellName) {
            if (worksheet.options.mergeCells?.[cellName]) {
                data = [worksheet.options.mergeCells[cellName][0], worksheet.options.mergeCells[cellName][1]];
            } else {
                data = null;
            }
        } else {
            if (worksheet.options.mergeCells) {
                for (const [key, [colspan, rowspan]] of Object.entries(worksheet.options.mergeCells)) {
                    data[key] = [colspan, rowspan];
                }
            }
        }

        return data;
    }

    /**
     * Merge cells.
     * @param cellName - The first cell of the range. Without it, the highlighted cells.
     * @param colspan - The number of columns.
     * @param rowspan - The number of rows.
     * @param ignoreHistoryAndEvents - Add no history and trigger no event.
     * @param keepSelection - Do not select the first cell of the range, as for the ranges of the options.
     */
    set(cellName?: string, colspan?: number, rowspan?: number, ignoreHistoryAndEvents?: boolean | 1, keepSelection?: boolean): null | undefined {
        const worksheet = this.worksheet;

        if (!cellName) {
            if (!worksheet.selection.highlighted.length) {
                alert(translate('No cells selected'));
                return null;
            } else {
                // The highlighted items are cell records, not td elements
                const first = worksheet.selection.highlighted[0].element;
                const last = worksheet.selection.highlighted[worksheet.selection.highlighted.length - 1].element;
                const x1 = parseInt(first.getAttribute('data-x'));
                const y1 = parseInt(first.getAttribute('data-y'));
                const x2 = parseInt(last.getAttribute('data-x'));
                const y2 = parseInt(last.getAttribute('data-y'));
                cellName = getColumnNameFromId([x1, y1]);
                colspan = x2 - x1 + 1;
                rowspan = y2 - y1 + 1;
            }
        } else if (typeof cellName !== 'string') {
            return null;
        }

        const test = this.check(cellName, colspan, rowspan);

        if (test) {
            alert(translate(test));
        } else {
            const cell = getIdFromColumnName(cellName);
            // A missing span becomes 1
            colspan = colspan! > 1 ? colspan! : 1;
            rowspan = rowspan! > 1 ? rowspan! : 1;
            const data = this.merge(cellName, colspan, rowspan);

            // The build of a worksheet with many merged ranges would select each range and dispatch onselection for it
            if (!keepSelection) {
                worksheet.selection.update(worksheet.records[cell[1]][cell[0]].element);
            }

            if (!ignoreHistoryAndEvents) {
                worksheet.history.add({
                    action: 'setMerge',
                    column: cellName,
                    colspan,
                    rowspan,
                    data,
                });

                worksheet.dispatch('onmerge', worksheet, { [cellName]: [colspan, rowspan] });
            }
        }
    }

    /**
     * Why a range cannot be merged, or false when it can. The first cell of a range of the options can be merged once.
     */
    private check(cellName: string, colspan: number | undefined, rowspan: number | undefined) {
        const worksheet = this.worksheet;

        const cell = getIdFromColumnName(cellName);

        if (worksheet.options.mergeCells?.[cellName]) {
            if (worksheet.records[cell[1]][cell[0]].element.getAttribute('data-merged')) {
                return 'Cell already merged';
            }
        } else if ((!colspan || colspan < 2) && (!rowspan || rowspan < 2)) {
            return 'Invalid merged properties';
        } else {
            // A missing span is one cell
            for (let j = cell[1]; j < cell[1] + (rowspan || 1); j++) {
                for (let i = cell[0]; i < cell[0] + (colspan || 1); i++) {
                    if (worksheet.records[j][i].element.getAttribute('data-merged')) {
                        return 'There is a conflict with another merged cell';
                    }
                }
            }
        }

        return false;
    }

    /**
     * Merge a range that check accepted, without a history entry, an event or a change of the selection. The other cells
     * of the range lose their values.
     * @returns The values of the other cells of the range, for the history.
     */
    private merge(cellName: string, colspan: number, rowspan: number) {
        const worksheet = this.worksheet;

        const cell = getIdFromColumnName(cellName);

        // Add property
        if (colspan > 1) {
            worksheet.records[cell[1]][cell[0]].element.setAttribute('colspan', colspan);
        }
        if (rowspan > 1) {
            worksheet.records[cell[1]][cell[0]].element.setAttribute('rowspan', rowspan);
        }
        // Keep links to the existing nodes
        if (!worksheet.options.mergeCells) {
            worksheet.options.mergeCells = {};
        }

        // The public typings have only colspan and rowspan
        (worksheet.options.mergeCells as Record<string, MergeInfo>)[cellName] = [colspan, rowspan, []];
        // Mark cell as merged
        worksheet.records[cell[1]][cell[0]].element.setAttribute('data-merged', 'true');
        // Overflow
        worksheet.records[cell[1]][cell[0]].element.style.overflow = 'hidden';
        // History data
        const data = [];
        // Adjust the nodes
        for (let y = cell[1]; y < cell[1] + rowspan; y++) {
            for (let x = cell[0]; x < cell[0] + colspan; x++) {
                if (!(cell[0] === x && cell[1] === y)) {
                    data.push(worksheet.options.data[y][x]);
                    worksheet.cells.update(x, y, '', true);
                    (worksheet.options.mergeCells[cellName] as MergeInfo)[2]!.push(worksheet.records[y][x].element);
                    worksheet.records[y][x].element.style.display = 'none';
                    worksheet.records[y][x].element = worksheet.records[cell[1]][cell[0]].element;
                }
            }
        }
        // A column of the range can be hidden
        this.updateFirstCell(cellName);

        return data;
    }

    /**
     * Split a merged range.
     * @param cellName - The first cell of the range.
     * @param data - The values of the other cells of the range, which the history kept.
     * @param keepOptions - Keep the range in the options.
     */
    remove(cellName: string, data?: jspreadsheet.CellValue[] | null, keepOptions?: boolean) {
        const worksheet = this.worksheet;

        if (worksheet.options.mergeCells?.[cellName]) {
            const beforeMerges = { [cellName]: worksheet.options.mergeCells[cellName] };
            const cell = getIdFromColumnName(cellName);
            const [colspan, rowspan] = worksheet.options.mergeCells[cellName];

            this.unmerge(cellName, data);

            // Update selection. The selection events still see the range.
            worksheet.selection.update(worksheet.records[cell[1]][cell[0]].element, worksheet.records[cell[1] + rowspan - 1][cell[0] + colspan - 1].element);

            this.forget(cellName, keepOptions);

            worksheet.dispatch('onunmerge', worksheet, cellName, beforeMerges);
        }
    }

    /**
     * Split a merged range in the table, without an event or a change of the selection. The range stays in the options
     * until forget removes it.
     * @param cellName - The first cell of the range.
     * @param data - The values of the other cells of the range, which the history kept.
     */
    private unmerge(cellName: string, data?: jspreadsheet.CellValue[] | null) {
        const worksheet = this.worksheet;

        const cell = getIdFromColumnName(cellName);
        worksheet.records[cell[1]][cell[0]].element.removeAttribute('colspan');
        worksheet.records[cell[1]][cell[0]].element.removeAttribute('rowspan');
        worksheet.records[cell[1]][cell[0]].element.removeAttribute('data-merged');
        // Each cell is visible when its column is visible
        worksheet.records[cell[1]][cell[0]].element.style.display = worksheet.headers[cell[0]].style.display === 'none' ? 'none' : '';
        const info = worksheet.options.mergeCells![cellName] as MergeInfo;

        let index = 0;

        for (let j = 0; j < info[1]; j++) {
            for (let i = 0; i < info[0]; i++) {
                if (j > 0 || i > 0) {
                    worksheet.records[cell[1] + j][cell[0] + i].element = info[2]![index];
                    worksheet.records[cell[1] + j][cell[0] + i].element.style.display = worksheet.headers[cell[0] + i].style.display === 'none' ? 'none' : '';
                    // Recover data
                    if (data?.[index]) {
                        worksheet.cells.update(cell[0] + i, cell[1] + j, data[index]);
                    }
                    index++;
                }
            }
        }
    }

    /**
     * After a split: remove the range from the options, and show a first cell in a column with the type hidden with the
     * options of its own column again.
     */
    private forget(cellName: string, keepOptions?: boolean) {
        const worksheet = this.worksheet;

        const cell = getIdFromColumnName(cellName);

        if (!keepOptions) {
            delete worksheet.options.mergeCells![cellName];
        }

        if (worksheet.options.columns?.[cell[0]]?.type === 'hidden') {
            worksheet.cells.render(cell[0], cell[1]);
        }
    }

    /**
     * Make the merged ranges equal to merges. The ranges that do not change stay. No dialog, history entry or change of the
     * selection. The method checks each entry before it changes a range, because the entries can come from a server.
     * @returns The first cells of the ranges that cannot be merged.
     */
    replace(merges: Record<string, [number, number]>) {
        const worksheet = this.worksheet;

        if (!merges || typeof merges !== 'object' || Array.isArray(merges)) {
            throw new TypeError('replaceMerges needs an object of ranges');
        }

        // A range of the options that the build did not merge has no list of cells: it is not merged
        const current = (worksheet.options.mergeCells || {}) as Record<string, MergeInfo>;
        for (const cellName of Object.keys(current)) {
            if (!Array.isArray(current[cellName][2])) {
                delete current[cellName];
            }
        }

        const refused: string[] = [];
        const wanted: [string, number, number][] = [];
        for (const [cellName, spans] of Object.entries(merges as Record<string, unknown>)) {
            const valid =
                /^[A-Z]+[1-9][0-9]*$/.test(cellName) &&
                Array.isArray(spans) &&
                spans.length === 2 &&
                spans.every((span) => Number.isInteger(span) && span > 0) &&
                spans[0] * spans[1] > 1;
            if (valid) {
                wanted.push([cellName, spans[0], spans[1]]);
            } else {
                refused.push(cellName);
            }
        }

        const removed: string[] = [];
        for (const cellName of Object.keys(current)) {
            const target = wanted.find(([name]) => name === cellName);
            if (target?.[1] !== current[cellName][0] || target[2] !== current[cellName][1]) {
                removed.push(cellName);
            }
        }
        const beforeMerges = removed.map((cellName) => ({ [cellName]: [current[cellName][0], current[cellName][1]] as [number, number] }));
        for (const cellName of removed) {
            this.unmerge(cellName);
            this.forget(cellName);
        }

        const added: Record<string, [number, number]> = {};
        for (const [cellName, colspan, rowspan] of wanted) {
            if (Object.hasOwn(current, cellName)) {
                continue;
            }
            const cell = getIdFromColumnName(cellName);
            if (cell[1] + rowspan > worksheet.records.length || cell[0] + colspan > worksheet.headers.length || this.check(cellName, colspan, rowspan)) {
                refused.push(cellName);
                continue;
            }

            // The editor of a cell that the range hides loses its cell, so it closes without a save
            const editor = worksheet.editor;
            if (
                editor.isOpen() &&
                editor.x >= cell[0] &&
                editor.x < cell[0] + colspan &&
                editor.y >= cell[1] &&
                editor.y < cell[1] + rowspan &&
                (editor.x !== cell[0] || editor.y !== cell[1])
            ) {
                worksheet.closeEditor(editor.cell, false);
            }

            this.merge(cellName, colspan, rowspan);
            added[cellName] = [colspan, rowspan];
        }

        // The tab stop can be in a cell that a new range hides
        if (removed.length || Object.keys(added).length) {
            worksheet.selection.updateTabStop();
        }

        removed.forEach((cellName, k) => {
            worksheet.dispatch('onunmerge', worksheet, cellName, beforeMerges[k]);
        });
        if (Object.keys(added).length) {
            worksheet.dispatch('onmerge', worksheet, added);
        }

        return refused;
    }

    /**
     * Remove all merged cells
     */
    destroy(keepOptions?: boolean) {
        const worksheet = this.worksheet;

        // Remove any merged cells
        if (worksheet.options.mergeCells) {
            for (const key of Object.keys(worksheet.options.mergeCells)) {
                this.remove(key, null, keepOptions);
            }
        }
    }
}
