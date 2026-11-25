import jSuites from 'jsuites';

import { parseCSV } from './helpers.js';
import dispatch from './dispatch.js';
import { updateCell } from './render.js';
import { updateFormulaChain, updateTableReferences } from './references.js';
import { updateTable } from './table.js';
import { downGet, rightGet } from './keys.js';
import { updateSelectionFromCoords } from './selection.js';
import { getColumnNameFromId } from './internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, WorksheetMethod } from '../types/internal';

/**
 * Put a text in double quotes when it has the delimiter, a comma, a line break or a double quote, and double the double
 * quotes in it. A value that is not text stays as it is.
 * @param div - The delimiter, as a regular expression.
 */
const quote = function <T>(value: T, div: RegExp): T | string {
    // Only text has match
    const text = value as unknown as string;
    if (text.match && (text.match(div) || text.match(/,/g) || text.match(/\n/) || text.match(/"/))) {
        return '"' + text.replace(new RegExp('"', 'g'), '""') + '"';
    }
    return value;
};

/**
 * The rows of the nested headers, as text, each with the title of a header followed by an empty column for each column
 * that it spans.
 */
const getNestedHeadersText = function (this: Worksheet, delimiter: string) {
    const obj = this;

    let nestedHeaders = '';

    if (obj.options.nestedHeaders && obj.options.nestedHeaders.length > 0) {
        const headerRows = obj.options.nestedHeaders;

        for (let j = 0; j < headerRows.length; j++) {
            const nested = [];
            for (let i = 0; i < headerRows[j].length; i++) {
                const colspan = parseInt(headerRows[j][i].colspan);
                nested.push(headerRows[j][i].title);
                for (let c = 0; c < colspan - 1; c++) {
                    nested.push('');
                }
            }
            nestedHeaders += nested.join(delimiter) + '\r\n';
        }
    }

    return nestedHeaders;
};

/**
 * The rows of the copied cells: the data and the shown values as text, and the styles of the cells.
 * @param highlighted - Copy only the highlighted cells.
 * @param div - The delimiter, as a regular expression.
 * @param copyHeader - Put the headers before the first row.
 */
const getCopiedRows = function (this: Worksheet, highlighted: boolean, delimiter: string, div: RegExp, copyHeader: boolean) {
    const obj = this;

    const header = [];
    let col = [];
    let colLabel = [];
    const row = [];
    const rowLabel = [];
    const x = obj.options.data[0].length;
    const y = obj.options.data.length;
    let numOfCols = 0;
    let numOfRows = 0;

    // The styles of the copied cells, which a paste of the same text applies
    const style: string[] = [];

    // Go through the columns to get the data
    for (let j = 0; j < y; j++) {
        col = [];
        colLabel = [];

        for (let i = 0; i < x; i++) {
            // If cell is highlighted
            if (!highlighted || obj.records[j][i].element.classList.contains('highlight')) {
                if (copyHeader == true) {
                    header.push(obj.headers[i].textContent);
                }
                // Values
                const value = quote(obj.options.data[j][i], div);
                col.push(value);

                // Labels
                let label;

                if (obj.options.columns && obj.options.columns[i] && (obj.options.columns[i].type == 'checkbox' || obj.options.columns[i].type == 'radio')) {
                    label = value;
                } else {
                    label = quote(obj.records[j][i].element.innerHTML, div);
                }
                colLabel.push(label);

                // Get style
                // createCell gives each cell a style attribute
                const cellStyle = obj.records[j][i].element.getAttribute('style')!.replace('display: none;', '');
                style.push(cellStyle ? cellStyle : '');
            }
        }

        if (col.length) {
            if (copyHeader) {
                numOfCols = col.length;
                row.push(header.join(delimiter));
            }
            row.push(col.join(delimiter));
        }
        if (colLabel.length) {
            numOfRows++;
            if (copyHeader) {
                rowLabel.push(header.join(delimiter));
                copyHeader = false;
            }
            rowLabel.push(colLabel.join(delimiter));
        }
    }

    return { row, rowLabel, style, numOfCols, numOfRows };
};

/**
 * Copy method
 *
 * @param bool highlighted - Get only highlighted cells
 * @param delimiter - \t default to keep compatibility with excel
 * @return string value
 */
export const copy = function (
    this: Worksheet,
    highlighted: boolean,
    delimiter?: string,
    returnData?: boolean,
    includeHeaders?: boolean,
    download?: boolean,
    isCut?: boolean,
    processed?: boolean
) {
    const obj = this;

    if (!delimiter) {
        delimiter = '\t';
    }

    const div = new RegExp(delimiter, 'ig');

    // Controls
    const x = obj.options.data[0].length;
    const y = obj.options.data.length;
    let copyHeader = false;
    let headers = '';
    let nestedHeaders = '';

    if (download && (obj.parent.config.includeHeadersOnDownload == true || includeHeaders)) {
        // Nested headers
        nestedHeaders = getNestedHeadersText.call(obj, delimiter);

        copyHeader = true;
    }

    const { row, rowLabel, style, numOfCols, numOfRows } = getCopiedRows.call(obj, highlighted, delimiter, div, copyHeader);

    if (x == numOfCols && y == numOfRows) {
        headers = nestedHeaders;
    }

    // Final string
    const str = headers + row.join('\r\n');
    let strLabel = headers + rowLabel.join('\r\n');

    // Create a hidden textarea to copy the values
    if (!returnData) {
        // Without a selection, no cell is highlighted: copy nothing and keep the clipboard
        if (!obj.selection.range) {
            return;
        }

        // Paste event
        const selectedRange: [number, number, number, number] = [
            Math.min(obj.selection.range[0], obj.selection.range[2]),
            Math.min(obj.selection.range[1], obj.selection.range[3]),
            Math.max(obj.selection.range[0], obj.selection.range[2]),
            Math.max(obj.selection.range[1], obj.selection.range[3]),
        ];

        const result = dispatch.call(obj, 'oncopy', obj, selectedRange, strLabel, isCut);

        if (result) {
            strLabel = result;
        } else if (result === false) {
            return false;
        }

        obj.clipboard.write(strLabel);
    }

    const data = processed == true ? strLabel : str;

    // A copy to the clipboard keeps the data, so that a paste of the same text gets the styles too. A cancelled copy or a
    // download keeps what was copied before.
    if (!returnData) {
        obj.clipboard.keep(data, style);

        // Any exiting border should go
        obj.clipboard.clearBorder();

        // Border
        if (obj.selection.highlighted) {
            obj.clipboard.showBorder(obj.selection.highlighted);
        }
    }

    return data;
};

/**
 * Repeat the rows and the styles of a paste to fill the selection, when the selection is a multiple of their width or
 * their height.
 * @param w - The width of the selection.
 * @param h - The height of the selection.
 */
const repeatToSelection = function (rows: jspreadsheet.CellValue[][], style: string[] | null, w: number, h: number) {
    const srcW = rows[0].length;
    if (w > 1 && Number.isInteger(w / srcW)) {
        const repeats = w / srcW;
        if (style) {
            const newStyle = [];
            for (let i = 0; i < style.length; i += srcW) {
                const chunk = style.slice(i, i + srcW);
                for (let j = 0; j < repeats; j++) {
                    newStyle.push(...chunk);
                }
            }
            style = newStyle;
        }

        const arrayB = rows.map(function (row, i) {
            const arrayC = Array.from({ length: repeats * row.length }, function (e: unknown, i: number) {
                return row[i % row.length];
            });
            return arrayC;
        });
        rows = arrayB;
    }
    const srcH = rows.length;
    if (h > 1 && Number.isInteger(h / srcH)) {
        const repeats = h / srcH;
        if (style) {
            const newStyle = [];
            for (let j = 0; j < repeats; j++) {
                newStyle.push(...style);
            }
            style = newStyle;
        }
        const arrayB = Array.from({ length: repeats * srcH }, function (e: unknown, i: number) {
            return rows[i % srcH];
        });
        rows = arrayB;
    }

    return { rows, style };
};

/**
 * Insert the columns and the rows that the paste needs after the target cell, counting the hidden ones, and then update
 * the references once.
 */
const insertSpaceForPaste = function (this: Worksheet, colIndex: number, rowIndex: number, rows: jspreadsheet.CellValue[][]) {
    const obj = this;

    const hiddenColCount = obj.headers.slice(colIndex).filter((x) => x.style.display === 'none').length;
    const expandedColCount = colIndex + hiddenColCount + rows[0].length;
    const currentColCount = obj.headers.length;
    if (expandedColCount > currentColCount) {
        obj.skipUpdateTableReferences = true;
        obj.insertColumn(expandedColCount - currentColCount);
    }
    const hiddenRowCount = obj.rows.slice(rowIndex).filter((x) => x.element.style.display === 'none').length;
    const expandedRowCount = rowIndex + hiddenRowCount + rows.length;
    const currentRowCount = obj.rows.length;
    if (expandedRowCount > currentRowCount) {
        obj.skipUpdateTableReferences = true;
        obj.insertRow(expandedRowCount - currentRowCount);
    }

    if (obj.skipUpdateTableReferences) {
        obj.skipUpdateTableReferences = false;
        updateTableReferences.call(obj);
    }
};

/**
 * Write the rows in the cells from the target cell, skipping the hidden cells, with the styles of the copied cells.
 * Insert a column or a row when the data goes past the last one, if the worksheet allows it.
 * @returns The records for the history, the new and the old styles, and the last cell written.
 */
const pasteRows = function (this: Worksheet, x: number, y: number, rows: jspreadsheet.CellValue[][], style: string[] | null) {
    const obj = this;

    // Records
    let i = 0;
    let j = 0;
    const records = [];
    const newStyle: Record<string, string> = {};
    const oldStyle: Record<string, string> = {};
    let styleIndex = 0;

    // Index
    let colIndex = parseInt(x);
    let rowIndex = parseInt(y);
    let row = null;

    while ((row = rows[j])) {
        i = 0;
        colIndex = parseInt(x);

        while (row[i] != null) {
            let value = row[i];

            if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].type == 'calendar') {
                // The jsuites typings miss extractDateFromString
                value = jSuites.calendar.extractDateFromString(
                    value,
                    (obj.options.columns[i].options && obj.options.columns[i].options.format) || 'YYYY-MM-DD'
                );
            }

            // Update and keep history
            const record = updateCell.call(obj, colIndex, rowIndex, value);
            // Keep history
            records.push(record);
            // Update all formulas in the chain
            updateFormulaChain.call(obj, colIndex, rowIndex, records);
            // Style
            if (style && style[styleIndex]) {
                const columnName = getColumnNameFromId([colIndex, rowIndex]);
                newStyle[columnName] = style[styleIndex];
                // The style text of one cell
                oldStyle[columnName] = obj.getStyle(columnName) as string;
                obj.records[rowIndex][colIndex].element.setAttribute('style', style[styleIndex]);
                styleIndex++;
            }
            i++;
            if (row[i] != null) {
                if (colIndex >= obj.headers.length - 1) {
                    // If the pasted column is out of range, create it if possible
                    if (obj.options.allowInsertColumn != false) {
                        obj.insertColumn();
                        // Otherwise skip the pasted data that overflows
                    } else {
                        break;
                    }
                }
                colIndex = rightGet.call(obj, colIndex, rowIndex);
            }
        }

        j++;
        if (rows[j]) {
            if (rowIndex >= obj.rows.length - 1) {
                // If the pasted row is out of range, create it if possible
                if (obj.options.allowInsertRow != false) {
                    obj.insertRow();
                    // Otherwise skip the pasted data that overflows
                } else {
                    break;
                }
            }
            rowIndex = downGet.call(obj, x, rowIndex);
        }
    }

    return { records, newStyle, oldStyle, lastColumn: colIndex, lastRow: rowIndex };
};

/**
 * Jspreadsheet paste method
 *
 * @param x target column
 * @param y target row
 * @param data paste data. if data hash is the same as the copied data, apply style from copied cells
 * @return string value
 */
export const paste: WorksheetMethod<'paste'> = function (x, y, data) {
    const obj = this;

    // Controls
    const isCopy = obj.clipboard.matches(data);
    let style = isCopy ? obj.clipboard.style : null;

    // Depending on the behavior
    if (isCopy) {
        data = obj.clipboard.data;
    }

    // Split new line. The rows and columns of the text, which can grow below to fill the selection, or be replaced by onbeforepaste
    let rows: jspreadsheet.CellValue[][] = parseCSV(data, '\t');

    // The end of the selection, to repeat the data over it. Without a selection, the data is pasted once.
    const ex = obj.selection.range ? obj.selection.range[2] : x;
    const ey = obj.selection.range ? obj.selection.range[3] : y;

    const w = ex - x + 1;
    const h = ey - y + 1;

    // Modify data to allow wor extending paste range in multiples of input range
    ({ rows, style } = repeatToSelection(rows, style, w, h));

    // Paste filter
    const ret = dispatch.call(
        obj,
        'onbeforepaste',
        obj,
        rows.map(function (row) {
            return row.map(function (item) {
                return { value: item };
            });
        }),
        x,
        y
    );

    if (ret === false) {
        return false;
    } else if (typeof ret == 'string') {
        // The text to paste instead
        rows = parseCSV(ret, '\t');
    } else if (ret && ret !== true) {
        // Rows of values to paste instead, which the typings do not declare
        rows = ret;
    }

    if (x != null && y != null && rows) {
        // Index
        const colIndex = parseInt(x);
        const rowIndex = parseInt(y);

        insertSpaceForPaste.call(obj, colIndex, rowIndex, rows);

        // Go through the columns to get the data
        const { records, newStyle, oldStyle, lastColumn, lastRow } = pasteRows.call(obj, x, y, rows, style);

        // Select the new cells
        updateSelectionFromCoords.call(obj, x, y, lastColumn, lastRow);

        // Update history
        obj.history.add({
            action: 'setValue',
            records: records,
            selection: obj.selection.range && [...obj.selection.range],
            newStyle: newStyle,
            oldStyle: oldStyle,
        });

        // Update table
        updateTable.call(obj);

        // Paste event
        const eventRecords = [];

        for (let j = 0; j < rows.length; j++) {
            for (let i = 0; i < rows[j].length; i++) {
                eventRecords.push({
                    x: i + x,
                    y: j + y,
                    value: rows[j][i],
                });
            }
        }

        dispatch.call(obj, 'onpaste', obj, eventRecords);

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
        dispatch.call(obj, 'onafterchanges', obj, onafterchangesRecords as unknown as jspreadsheet.CellChange[]);
    }

    obj.clipboard.clearBorder();
};
