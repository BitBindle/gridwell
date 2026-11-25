import jSuites from 'jsuites';

import { parseCSV } from './helpers.js';
import dispatch from './dispatch.js';
import { setHistory } from './history.js';
import { updateCell, updateFormulaChain, updateTable, updateTableReferences } from './internal.js';
import { downGet, rightGet } from './keys.js';
import { hash, removeCopyingSelection, updateSelectionFromCoords } from './selection.js';
import { getColumnNameFromId } from './internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, WorksheetMethod } from '../types/internal';

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
    const header = [];
    let col = [];
    let colLabel = [];
    const row = [];
    const rowLabel = [];
    const x = obj.options.data[0].length;
    const y = obj.options.data.length;
    let copyHeader = false;
    let headers = '';
    let nestedHeaders = '';
    let numOfCols = 0;
    let numOfRows = 0;

    // Partial copy
    let copyX = 0;
    let copyY = 0;
    let isPartialCopy = true;
    // Go through the columns to get the data
    for (let j = 0; j < y; j++) {
        for (let i = 0; i < x; i++) {
            // If cell is highlighted
            if (!highlighted || obj.records[j][i].element.classList.contains('highlight')) {
                if (copyX <= i) {
                    copyX = i;
                }
                if (copyY <= j) {
                    copyY = j;
                }
            }
        }
    }
    if (x === copyX + 1 && y === copyY + 1) {
        isPartialCopy = false;
    }

    if (download && (obj.parent.config.includeHeadersOnDownload == true || includeHeaders)) {
        // Nested headers
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

        copyHeader = true;
    }

    // Reset container
    obj.style = [];

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
                let value = obj.options.data[j][i];
                // Only text has match
                if (
                    (value as string).match &&
                    ((value as string).match(div) || (value as string).match(/,/g) || (value as string).match(/\n/) || (value as string).match(/"/))
                ) {
                    value = (value as string).replace(new RegExp('"', 'g'), '""');
                    value = '"' + value + '"';
                }
                col.push(value);

                // Labels
                let label;

                if (obj.options.columns && obj.options.columns[i] && (obj.options.columns[i].type == 'checkbox' || obj.options.columns[i].type == 'radio')) {
                    label = value;
                } else {
                    label = obj.records[j][i].element.innerHTML;
                    if (label.match && (label.match(div) || label.match(/,/g) || label.match(/\n/) || label.match(/"/))) {
                        // Scape double quotes
                        label = label.replace(new RegExp('"', 'g'), '""');
                        label = '"' + label + '"';
                    }
                }
                colLabel.push(label);

                // Get style
                // createCell gives each cell a style attribute
                const cellStyle = obj.records[j][i].element.getAttribute('style')!.replace('display: none;', '');
                obj.style.push(cellStyle ? cellStyle : '');
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

    if (x == numOfCols && y == numOfRows) {
        headers = nestedHeaders;
    }

    // Final string
    const str = headers + row.join('\r\n');
    let strLabel = headers + rowLabel.join('\r\n');

    // Create a hidden textarea to copy the values
    if (!returnData) {
        // Without a selection, no cell is highlighted: copy nothing and keep the clipboard
        if (!obj.selectedCell) {
            return;
        }

        // Paste event
        const selectedRange: [number, number, number, number] = [
            Math.min(obj.selectedCell[0], obj.selectedCell[2]),
            Math.min(obj.selectedCell[1], obj.selectedCell[3]),
            Math.max(obj.selectedCell[0], obj.selectedCell[2]),
            Math.max(obj.selectedCell[1], obj.selectedCell[3]),
        ];

        const result = dispatch.call(obj, 'oncopy', obj, selectedRange, strLabel, isCut);

        if (result) {
            strLabel = result;
        } else if (result === false) {
            return false;
        }

        obj.textarea.value = strLabel;
        obj.textarea.select();
        document.execCommand('copy');
    }

    // Keep data
    if (processed == true) {
        obj.data = strLabel;
    } else {
        obj.data = str;
    }
    // Keep non visible information
    obj.hashString = hash.call(obj, obj.data);

    // Any exiting border should go
    if (!returnData) {
        removeCopyingSelection.call(obj);

        // Border
        if (obj.highlighted) {
            for (let i = 0; i < obj.highlighted.length; i++) {
                obj.highlighted[i].element.classList.add('copying');
                if (obj.highlighted[i].element.classList.contains('highlight-left')) {
                    obj.highlighted[i].element.classList.add('copying-left');
                }
                if (obj.highlighted[i].element.classList.contains('highlight-right')) {
                    obj.highlighted[i].element.classList.add('copying-right');
                }
                if (obj.highlighted[i].element.classList.contains('highlight-top')) {
                    obj.highlighted[i].element.classList.add('copying-top');
                }
                if (obj.highlighted[i].element.classList.contains('highlight-bottom')) {
                    obj.highlighted[i].element.classList.add('copying-bottom');
                }
            }
        }
    }

    return obj.data;
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
    const dataHash = hash(data);
    let style = dataHash == obj.hashString ? obj.style : null;

    // Depending on the behavior
    if (dataHash == obj.hashString) {
        data = obj.data;
    }

    // Split new line. The rows and columns of the text, which can grow below to fill the selection, or be replaced by onbeforepaste
    let rows: jspreadsheet.CellValue[][] = parseCSV(data, '\t');

    // The end of the selection, to repeat the data over it. Without a selection, the data is pasted once.
    const ex = obj.selectedCell ? obj.selectedCell[2] : x;
    const ey = obj.selectedCell ? obj.selectedCell[3] : y;

    const w = ex - x + 1;
    const h = ey - y + 1;

    // Modify data to allow wor extending paste range in multiples of input range
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

        // Go through the columns to get the data
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

        // Select the new cells
        updateSelectionFromCoords.call(obj, x, y, colIndex, rowIndex);

        // Update history
        setHistory.call(obj, {
            action: 'setValue',
            records: records,
            selection: obj.selectedCell,
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

    removeCopyingSelection();
};
