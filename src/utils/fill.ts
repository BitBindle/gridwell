import dispatch from './dispatch.js';
import { updateCell, updateFormula, updateFormulaChain, updateTable } from './internal.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import type { Worksheet } from '../types/internal';
import type { CellChange, CellValue } from '../jspreadsheet';

// The fill of cells with a drag of the selection corner. It writes cells, so it is apart from the selection, which the
// modules that write cells import.

const doubleDigitFormat = function (v: number | string) {
    v = '' + v;
    if (v.length == 1) {
        v = '0' + v;
    }
    return v;
};

/**
 * Helper function to copy data using the corner icon
 */
export const copyData = function (this: Worksheet, o: HTMLElement, d: HTMLElement) {
    const obj = this;

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
                                    value = updateFormula(value as string, affectedTokens);
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
                updateFormulaChain.call(obj, i, j, records);
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
};
