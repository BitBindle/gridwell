import jSuites from 'jsuites';
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
            const keys = Object.keys(worksheet.options.mergeCells);
            for (let i = 0; i < keys.length; i++) {
                const info = getIdFromColumnName(keys[i], true);
                const colspan = worksheet.options.mergeCells[keys[i]][0];
                const x1 = info[0];
                const x2 = info[0] + (colspan > 1 ? colspan - 1 : 0);

                if (insertBefore == null) {
                    if (x1 <= x && x2 >= x) {
                        cols.push(keys[i]);
                    }
                } else {
                    if (insertBefore) {
                        if (x1 < x && x2 >= x) {
                            cols.push(keys[i]);
                        }
                    } else {
                        if (x1 <= x && x2 > x) {
                            cols.push(keys[i]);
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
            const keys = Object.keys(worksheet.options.mergeCells);
            for (let i = 0; i < keys.length; i++) {
                const info = getIdFromColumnName(keys[i], true);
                const rowspan = worksheet.options.mergeCells[keys[i]][1];
                const y1 = info[1];
                const y2 = info[1] + (rowspan > 1 ? rowspan - 1 : 0);

                if (insertBefore == null) {
                    if (y1 <= y && y2 >= y) {
                        rows.push(keys[i]);
                    }
                } else {
                    if (insertBefore) {
                        if (y1 < y && y2 >= y) {
                            rows.push(keys[i]);
                        }
                    } else {
                        if (y1 <= y && y2 > y) {
                            rows.push(keys[i]);
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

        const cell = getIdFromColumnName(cellName, true);
        const colspan = worksheet.options.mergeCells![cellName][0];

        let visible = 0;
        for (let x = cell[0]; x < cell[0] + colspan; x++) {
            if (worksheet.headers[x].style.display != 'none') {
                visible++;
            }
        }

        const element = worksheet.records[cell[1]][cell[0]].element;
        if (visible == 0) {
            element.style.display = 'none';
        } else {
            element.style.display = '';
            if (colspan > 1) {
                element.setAttribute('colspan', visible);
            }
        }

        // A first cell in a column with the type hidden shows the range with the options of another column
        if (worksheet.options.columns && worksheet.options.columns[cell[0]] && worksheet.options.columns[cell[0]].type == 'hidden') {
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

        const keys = Object.keys(worksheet.options.mergeCells);
        for (let i = 0; i < keys.length; i++) {
            const cell = getIdFromColumnName(keys[i], true);
            const colspan = worksheet.options.mergeCells[keys[i]][0];

            for (let j = 0; j < columns.length; j++) {
                const x = parseInt(columns[j]);
                if (x >= cell[0] && x < cell[0] + colspan) {
                    this.updateFirstCell(keys[i]);
                    break;
                }
            }
        }
    }

    /**
     * Merge cells
     * @param cellName
     * @param colspan
     * @param rowspan
     * @param ignoreHistoryAndEvents
     */
    get(cellName?: string) {
        const worksheet = this.worksheet;

        let data: Record<string, [number, number]> | [number, number] | null = {};
        if (cellName) {
            if (worksheet.options.mergeCells && worksheet.options.mergeCells[cellName]) {
                data = [worksheet.options.mergeCells[cellName][0], worksheet.options.mergeCells[cellName][1]];
            } else {
                data = null;
            }
        } else {
            if (worksheet.options.mergeCells) {
                const keys = Object.keys(worksheet.options.mergeCells);
                for (let i = 0; i < keys.length; i++) {
                    data[keys[i]] = [worksheet.options.mergeCells[keys[i]][0], worksheet.options.mergeCells[keys[i]][1]];
                }
            }
        }

        return data;
    }

    /**
     * Merge cells
     * @param cellName
     * @param colspan
     * @param rowspan
     * @param ignoreHistoryAndEvents
     */
    set(cellName?: string, colspan?: number, rowspan?: number, ignoreHistoryAndEvents?: boolean | 1): null | undefined {
        const worksheet = this.worksheet;

        let test: string | false = false;

        if (!cellName) {
            if (!worksheet.selection.highlighted.length) {
                alert(jSuites.translate('No cells selected'));
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

        const cell = getIdFromColumnName(cellName, true);

        if (worksheet.options.mergeCells && worksheet.options.mergeCells[cellName]) {
            if (worksheet.records[cell[1]][cell[0]].element.getAttribute('data-merged')) {
                test = 'Cell already merged';
            }
        } else if ((!colspan || colspan < 2) && (!rowspan || rowspan < 2)) {
            test = 'Invalid merged properties';
        } else {
            // A missing span is one cell
            for (let j = cell[1]; j < cell[1] + (rowspan || 1); j++) {
                for (let i = cell[0]; i < cell[0] + (colspan || 1); i++) {
                    if (worksheet.records[j][i].element.getAttribute('data-merged')) {
                        test = 'There is a conflict with another merged cell';
                    }
                }
            }
        }

        if (test) {
            alert(jSuites.translate(test));
        } else {
            // Add property. A missing span becomes 1, so from here on the spans are numbers.
            if (colspan! > 1) {
                worksheet.records[cell[1]][cell[0]].element.setAttribute('colspan', colspan!);
            } else {
                colspan = 1;
            }
            if (rowspan! > 1) {
                worksheet.records[cell[1]][cell[0]].element.setAttribute('rowspan', rowspan!);
            } else {
                rowspan = 1;
            }
            // Keep links to the existing nodes
            if (!worksheet.options.mergeCells) {
                worksheet.options.mergeCells = {};
            }

            // The public typings have only colspan and rowspan
            (worksheet.options.mergeCells as Record<string, MergeInfo>)[cellName] = [colspan!, rowspan!, []];
            // Mark cell as merged
            worksheet.records[cell[1]][cell[0]].element.setAttribute('data-merged', 'true');
            // Overflow
            worksheet.records[cell[1]][cell[0]].element.style.overflow = 'hidden';
            // History data
            const data = [];
            // Adjust the nodes
            for (let y = cell[1]; y < cell[1] + rowspan!; y++) {
                for (let x = cell[0]; x < cell[0] + colspan!; x++) {
                    if (!(cell[0] == x && cell[1] == y)) {
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
            // In the initialization is not necessary keep the history
            worksheet.selection.update(worksheet.records[cell[1]][cell[0]].element);

            if (!ignoreHistoryAndEvents) {
                worksheet.history.add({
                    action: 'setMerge',
                    column: cellName,
                    colspan: colspan!,
                    rowspan: rowspan!,
                    data,
                });

                worksheet.dispatch('onmerge', worksheet, { [cellName]: [colspan!, rowspan!] });
            }
        }
    }

    /**
     * Remove merge by cellname
     * @param cellName
     */
    remove(cellName: string, data?: jspreadsheet.CellValue[] | null, keepOptions?: boolean) {
        const worksheet = this.worksheet;

        if (worksheet.options.mergeCells && worksheet.options.mergeCells[cellName]) {
            const beforeMerges = { [cellName]: worksheet.options.mergeCells[cellName] };
            const cell = getIdFromColumnName(cellName, true);
            worksheet.records[cell[1]][cell[0]].element.removeAttribute('colspan');
            worksheet.records[cell[1]][cell[0]].element.removeAttribute('rowspan');
            worksheet.records[cell[1]][cell[0]].element.removeAttribute('data-merged');
            // Each cell is visible when its column is visible
            worksheet.records[cell[1]][cell[0]].element.style.display = worksheet.headers[cell[0]].style.display == 'none' ? 'none' : '';
            const info = worksheet.options.mergeCells[cellName] as MergeInfo;

            let index = 0;

            let j, i;

            for (j = 0; j < info[1]; j++) {
                for (i = 0; i < info[0]; i++) {
                    if (j > 0 || i > 0) {
                        worksheet.records[cell[1] + j][cell[0] + i].element = info[2]![index];
                        worksheet.records[cell[1] + j][cell[0] + i].element.style.display =
                            worksheet.headers[cell[0] + i].style.display == 'none' ? 'none' : '';
                        // Recover data
                        if (data && data[index]) {
                            worksheet.cells.update(cell[0] + i, cell[1] + j, data[index]);
                        }
                        index++;
                    }
                }
            }

            // Update selection
            worksheet.selection.update(worksheet.records[cell[1]][cell[0]].element, worksheet.records[cell[1] + j - 1][cell[0] + i! - 1].element);

            if (!keepOptions) {
                delete worksheet.options.mergeCells[cellName];
            }

            // A first cell in a column with the type hidden uses the options of its own column again
            if (worksheet.options.columns && worksheet.options.columns[cell[0]] && worksheet.options.columns[cell[0]].type == 'hidden') {
                worksheet.cells.render(cell[0], cell[1]);
            }

            worksheet.dispatch('onunmerge', worksheet, cellName, beforeMerges);
        }
    }

    /**
     * Remove all merged cells
     */
    destroy(keepOptions?: boolean) {
        const worksheet = this.worksheet;

        // Remove any merged cells
        if (worksheet.options.mergeCells) {
            const keys = Object.keys(worksheet.options.mergeCells);
            for (let i = 0; i < keys.length; i++) {
                this.remove(keys[i], null, keepOptions);
            }
        }
    }
}
