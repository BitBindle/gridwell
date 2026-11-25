import dispatch from '../utils/dispatch.js';
import { updateCell } from '../utils/render.js';
import { updateTable } from '../utils/table.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type { DragColumnInfo, DragRowInfo, ResizeColumnInfo, ResizeRowInfo, CellChange, CellValue } from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

const doubleDigitFormat = function (v: number | string) {
    v = '' + v;
    if (v.length == 1) {
        v = '0' + v;
    }
    return v;
};

/**
 * The state of a mouse gesture on a worksheet: the column or row that the mouse resizes or moves, the drag of the
 * selection corner, and the element whose mouse cursor shows a resize. It fills the cells when a drag of the corner ends.
 *
 * Each worksheet has one, as worksheet.pointer. The mouse handlers in src/events/mouse.ts set it when a button goes down
 * and clear it when it goes up.
 */
export class Pointer {
    readonly worksheet: Worksheet;

    /**
     * The column or row that the mouse resizes, or null. Each member says that the other index is missing, so that a
     * test of column or row tells the type check which one it is.
     */
    resizing: null | (ResizeColumnInfo & { row?: undefined }) | (ResizeRowInfo & { column?: undefined }) = null;

    /** The column or row that the mouse moves, or null. */
    dragging: null | (DragColumnInfo & { row?: undefined }) | (DragRowInfo & { column?: undefined }) = null;

    /** The header or row element that shows a resize cursor, or null. */
    cursor: HTMLElement | null = null;

    /** True while the mouse drags the selection corner to fill cells. */
    corner = false;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Helper function to copy data using the corner icon
     */
    fill(o: HTMLElement, d: HTMLElement) {
        const obj = this.worksheet;

        // Get data from all selected cells
        const data = obj.getData(true, false);

        // Selected cells. The corner is only there when cells are selected.
        const h = obj.selection.bounds!;

        // Cells
        const x1 = parseInt(o.getAttribute('data-x'));
        const y1 = parseInt(o.getAttribute('data-y'));
        const x2 = parseInt(d.getAttribute('data-x'));
        const y2 = parseInt(d.getAttribute('data-y'));

        // Records
        const records = [];
        let breakControl = false;

        let rowNumber, colNumber;

        if (h[0] == x1) {
            // Vertical copy
            if (y1 < h[1]) {
                rowNumber = y1 - h[1];
            } else {
                rowNumber = 1;
            }
            colNumber = 0;
        } else {
            if (x1 < h[0]) {
                colNumber = x1 - h[0];
            } else {
                colNumber = 1;
            }
            rowNumber = 0;
        }

        // Copy data procedure
        let posx = 0;
        let posy = 0;

        for (let j = y1; j <= y2; j++) {
            // Skip hidden rows
            if (obj.rows[j] && obj.rows[j].element.style.display == 'none') {
                continue;
            }

            // Controls
            if (data[posy] == undefined) {
                posy = 0;
            }
            posx = 0;

            // Data columns
            if (h[0] != x1) {
                if (x1 < h[0]) {
                    colNumber = x1 - h[0];
                } else {
                    colNumber = 1;
                }
            }
            // Data columns
            for (let i = x1; i <= x2; i++) {
                // Update non-readonly
                if (
                    obj.records[j][i] &&
                    !obj.records[j][i].element.classList.contains('readonly') &&
                    obj.records[j][i].element.style.display != 'none' &&
                    breakControl == false
                ) {
                    // Stop if contains value
                    if (!obj.selection.fill.length) {
                        if (obj.options.data[j][i] != '') {
                            breakControl = true;
                            continue;
                        }
                    }

                    // Column
                    if (data[posy] == undefined) {
                        posx = 0;
                    } else if (data[posy][posx] == undefined) {
                        posx = 0;
                    }

                    // Value
                    let value: CellValue = data[posy][posx];

                    if (value && !data[1] && obj.parent.config.autoIncrement != false) {
                        if (
                            obj.options.columns &&
                            obj.options.columns[i] &&
                            (!obj.options.columns[i].type || obj.options.columns[i].type == 'text' || (obj.options.columns[i].type as string) == 'number')
                        ) {
                            if (('' + value).substr(0, 1) == '=') {
                                const tokens = (value as string).match(/([A-Z]+[0-9]+)/g);

                                if (tokens) {
                                    // An array used as a map, as before
                                    const affectedTokens = [] as unknown as Record<string, string>;
                                    for (let index = 0; index < tokens.length; index++) {
                                        const position = getIdFromColumnName(tokens[index], true);
                                        position[0] += colNumber;
                                        position[1] += rowNumber;
                                        if (position[1] < 0) {
                                            position[1] = 0;
                                        }
                                        const token = getColumnNameFromId([position[0], position[1]]);

                                        if (token != tokens[index]) {
                                            affectedTokens[tokens[index]] = token;
                                        }
                                    }
                                    // Update formula
                                    if (affectedTokens) {
                                        value = obj.formulas.updateFormula(value as string, affectedTokens);
                                    }
                                }
                            } else {
                                if (value == Number(value)) {
                                    value = Number(value) + rowNumber;
                                }
                            }
                        } else if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].type == 'calendar') {
                            // A calendar cell has a date as text
                            const date = new Date(value as string);
                            date.setDate(date.getDate() + rowNumber);
                            value =
                                date.getFullYear() +
                                '-' +
                                doubleDigitFormat(parseInt(date.getMonth() + 1)) +
                                '-' +
                                doubleDigitFormat(date.getDate()) +
                                ' ' +
                                '00:00:00';
                        }
                    }

                    records.push(updateCell.call(obj, i, j, value));

                    // Update all formulas in the chain
                    obj.formulas.updateChain(i, j, records);
                }
                posx++;
                if (h[0] != x1) {
                    colNumber++;
                }
            }
            posy++;
            rowNumber++;
        }

        // Update history
        obj.history.add({
            action: 'setValue',
            records: records,
            selection: obj.selection.range && [...obj.selection.range],
        });

        // Update table with custom configuration if applicable
        updateTable.call(obj);

        // On after changes
        const onafterchangesRecords = records.map(function (record) {
            return {
                x: record.x,
                y: record.y,
                value: record.value,
                oldValue: record.oldValue,
            };
        });

        // The coordinates can be text (issue #52), and a read-only cell has no value (issue #58)
        dispatch.call(obj, 'onafterchanges', obj, onafterchangesRecords as unknown as CellChange[]);
    }
}
