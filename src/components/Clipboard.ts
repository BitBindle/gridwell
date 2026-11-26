import jSuites from 'jsuites';
import { parseCSV } from '../utils/helpers.js';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * Put a text in double quotes when it has the delimiter, a comma, a line break or a double quote, and double the double
 * quotes in it. A value that is not text stays as it is.
 * @param div - The delimiter, as a regular expression.
 */
function quote<T>(value: T, div: RegExp): T | string {
    // Only text has match
    const text = value as unknown as string;
    if (text.match && (text.match(div) || text.match(/,/g) || text.match(/\n/) || text.match(/"/))) {
        return `"${text.replace(new RegExp('"', 'g'), '""')}"`;
    }
    return value;
}

/**
 * Repeat the rows and the styles of a paste to fill the selection, when the selection is a multiple of their width or
 * their height.
 * @param w - The width of the selection.
 * @param h - The height of the selection.
 */
function repeatToSelection<T>(rows: T[][], style: string[] | null, w: number, h: number) {
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

        const arrayB = rows.map((row, i) => {
            const arrayC = Array.from({ length: repeats * row.length }, (e: unknown, i: number) => {
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
        const arrayB = Array.from({ length: repeats * srcH }, (e: unknown, i: number) => {
            return rows[i % srcH];
        });
        rows = arrayB;
    }

    return { rows, style };
}

/**
 * The clipboard of a worksheet: the text it copied last, the styles of the copied cells, the hidden textarea that puts
 * text on the system clipboard, and the dashed border around the copied cells. It copies the selected cells, and pastes
 * text into the worksheet.
 *
 * Each worksheet has one, as worksheet.clipboard. copy, paste and download call it.
 */
export class Clipboard {
    readonly worksheet: Worksheet;

    /** The text of the last copy, or an empty string before the first copy. */
    data = '';

    /** The hash of data, or null before the first copy. A paste compares it with the hash of the pasted text. */
    hash: number | null = null;

    /** The styles of the copied cells, one entry for each cell, by row. */
    style: string[] = [];

    /** The hidden textarea that the copy selects, to put the text on the system clipboard. */
    textarea: HTMLTextAreaElement;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
        this.textarea = document.createElement('textarea');
        this.textarea.className = 'jss_textarea';
        this.textarea.tabIndex = -1;
        this.textarea.ariaHidden = 'true';
    }

    /** Keep the text of a copy and the styles of its cells. */
    keep(data: string, style: string[]) {
        this.data = data;
        this.style = style;
        this.hash = hashText(data);
    }

    /** Whether text is the text of the last copy, so that a paste of it can use the data and the styles of the copy. */
    matches(text: string) {
        return hashText(text) === this.hash;
    }

    /** Put text on the system clipboard. */
    write(text: string) {
        this.textarea.value = text;
        this.textarea.select();
        document.execCommand('copy');
    }

    /** Draw the dashed border around the copied cells, from the borders of their selection. */
    showBorder(cells: { element: HTMLTableCellElement }[]) {
        for (const cell of cells) {
            cell.element.classList.add('copying');
            if (cell.element.classList.contains('highlight-left')) {
                cell.element.classList.add('copying-left');
            }
            if (cell.element.classList.contains('highlight-right')) {
                cell.element.classList.add('copying-right');
            }
            if (cell.element.classList.contains('highlight-top')) {
                cell.element.classList.add('copying-top');
            }
            if (cell.element.classList.contains('highlight-bottom')) {
                cell.element.classList.add('copying-bottom');
            }
        }
    }

    /** Remove the dashed border from the cells of every worksheet in the document: there is one system clipboard. */
    clearBorder() {
        const copying = document.querySelectorAll('.jss_worksheet .copying');
        for (const cell of copying) {
            cell.classList.remove('copying');
            cell.classList.remove('copying-left');
            cell.classList.remove('copying-right');
            cell.classList.remove('copying-top');
            cell.classList.remove('copying-bottom');
        }
    }

    /**
     * The rows of the nested headers, as text, each with the title of a header followed by an empty column for each column
     * that it spans.
     */
    private getNestedHeadersText(delimiter: string) {
        const worksheet = this.worksheet;

        let nestedHeaders = '';

        if (worksheet.options.nestedHeaders && worksheet.options.nestedHeaders.length > 0) {
            const headerRows = worksheet.options.nestedHeaders;

            for (const headerRow of headerRows) {
                const nested = [];
                for (const header of headerRow) {
                    const colspan = parseInt(header.colspan);
                    nested.push(header.title);
                    for (let c = 0; c < colspan - 1; c++) {
                        nested.push('');
                    }
                }
                nestedHeaders += `${nested.join(delimiter)}\r\n`;
            }
        }

        return nestedHeaders;
    }

    /**
     * The rows of the copied cells: the data and the shown values as text, and the styles of the cells.
     * @param highlighted - Copy only the highlighted cells.
     * @param div - The delimiter, as a regular expression.
     * @param copyHeader - Put the headers before the first row.
     */
    private getCopiedRows(highlighted: boolean, delimiter: string, div: RegExp, copyHeader: boolean) {
        const worksheet = this.worksheet;

        const header = [];
        let col = [];
        let colLabel = [];
        const row = [];
        const rowLabel = [];
        const x = worksheet.options.data[0].length;
        const y = worksheet.options.data.length;
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
                if (!highlighted || worksheet.records[j][i].element.classList.contains('highlight')) {
                    if (copyHeader === true) {
                        header.push(worksheet.headers[i].textContent);
                    }
                    // Values
                    const value = quote(worksheet.options.data[j][i], div);
                    col.push(value);

                    // Labels
                    let label;

                    if (worksheet.options.columns?.[i] && (worksheet.options.columns[i].type === 'checkbox' || worksheet.options.columns[i].type === 'radio')) {
                        label = value;
                    } else {
                        label = quote(worksheet.records[j][i].element.innerHTML, div);
                    }
                    colLabel.push(label);

                    // Get style
                    // createCell gives each cell a style attribute
                    const cellStyle = worksheet.records[j][i].element.getAttribute('style')!.replace('display: none;', '');
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
    }

    /**
     * Copy the selected cells, or give them as text.
     * @param highlighted - Copy only the highlighted cells.
     * @param delimiter - The delimiter of the columns. A tab by default, as spreadsheet applications use.
     * @param returnData - Return the text instead of copying it.
     * @param includeHeaders - Put the headers before the first row.
     * @param download - The text is for a download.
     * @param isCut - The cells are cut.
     * @param processed - Copy the shown values instead of the data.
     */
    copy(highlighted: boolean, delimiter?: string, returnData?: boolean, includeHeaders?: boolean, download?: boolean, isCut?: boolean, processed?: boolean) {
        const worksheet = this.worksheet;

        if (!delimiter) {
            delimiter = '\t';
        }

        const div = new RegExp(delimiter, 'ig');

        // Controls
        const x = worksheet.options.data[0].length;
        const y = worksheet.options.data.length;
        let copyHeader = false;
        let headers = '';
        let nestedHeaders = '';

        if (download && (worksheet.parent.config.includeHeadersOnDownload === true || includeHeaders)) {
            // Nested headers
            nestedHeaders = this.getNestedHeadersText(delimiter);

            copyHeader = true;
        }

        const { row, rowLabel, style, numOfCols, numOfRows } = this.getCopiedRows(highlighted, delimiter, div, copyHeader);

        if (x === numOfCols && y === numOfRows) {
            headers = nestedHeaders;
        }

        // Final string
        const str = headers + row.join('\r\n');
        let strLabel = headers + rowLabel.join('\r\n');

        // Create a hidden textarea to copy the values
        if (!returnData) {
            // Without a selection, no cell is highlighted: copy nothing and keep the clipboard
            if (!worksheet.selection.range) {
                return;
            }

            // Paste event
            const selectedRange: [number, number, number, number] = [
                Math.min(worksheet.selection.range[0], worksheet.selection.range[2]),
                Math.min(worksheet.selection.range[1], worksheet.selection.range[3]),
                Math.max(worksheet.selection.range[0], worksheet.selection.range[2]),
                Math.max(worksheet.selection.range[1], worksheet.selection.range[3]),
            ];

            const result = worksheet.dispatch('oncopy', worksheet, selectedRange, strLabel, isCut);

            if (result) {
                strLabel = result;
            } else if (result === false) {
                return false;
            }

            worksheet.clipboard.write(strLabel);
        }

        const data = processed === true ? strLabel : str;

        // A copy to the clipboard keeps the data, so that a paste of the same text gets the styles too. A cancelled copy or a
        // download keeps what was copied before.
        if (!returnData) {
            worksheet.clipboard.keep(data, style);

            // Any exiting border should go
            worksheet.clipboard.clearBorder();

            // Border
            if (worksheet.selection.highlighted) {
                worksheet.clipboard.showBorder(worksheet.selection.highlighted);
            }
        }

        return data;
    }

    /**
     * Insert the columns and the rows that the paste needs after the target cell, counting the hidden ones, and then update
     * the references once.
     */
    private insertSpaceForPaste(colIndex: number, rowIndex: number, rows: jspreadsheet.CellValue[][]) {
        const worksheet = this.worksheet;

        const hiddenColCount = worksheet.headers.slice(colIndex).filter((x) => x.style.display === 'none').length;
        const expandedColCount = colIndex + hiddenColCount + rows[0].length;
        const currentColCount = worksheet.headers.length;
        if (expandedColCount > currentColCount) {
            worksheet.skipUpdateTableReferences = true;
            worksheet.insertColumn(expandedColCount - currentColCount);
        }
        const hiddenRowCount = worksheet.rows.slice(rowIndex).filter((x) => x.element.style.display === 'none').length;
        const expandedRowCount = rowIndex + hiddenRowCount + rows.length;
        const currentRowCount = worksheet.rows.length;
        if (expandedRowCount > currentRowCount) {
            worksheet.skipUpdateTableReferences = true;
            worksheet.insertRow(expandedRowCount - currentRowCount);
        }

        if (worksheet.skipUpdateTableReferences) {
            worksheet.skipUpdateTableReferences = false;
            worksheet.formulas.updateTableReferences();
        }
    }

    /**
     * Write the rows in the cells from the target cell, skipping the hidden cells, with the styles of the copied cells.
     * Insert a column or a row when the data goes past the last one, if the worksheet allows it.
     * @returns The records for the history, the new and the old styles, and the last cell written.
     */
    private pasteRows(x: number, y: number, rows: jspreadsheet.CellValue[][], style: string[] | null) {
        const worksheet = this.worksheet;

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

                // The options of the column that gets the value: for the first cell of a merged range over a hidden column,
                // the first visible column of the range
                const column = worksheet.options.columns?.[worksheet.cells.getColumnIndex(colIndex, rowIndex)];
                if (column?.type === 'calendar') {
                    // The jsuites typings miss extractDateFromString
                    value = jSuites.calendar.extractDateFromString(value, column.options?.format || 'YYYY-MM-DD');
                }

                // Update and keep history. A cell that did not change, such as a read-only cell, has no record.
                const record = worksheet.cells.update(colIndex, rowIndex, value);
                if (record) {
                    records.push(record);
                    // Update all formulas in the chain
                    worksheet.formulas.updateChain(colIndex, rowIndex, records);
                }
                // Style
                if (style?.[styleIndex]) {
                    const columnName = getColumnNameFromId([colIndex, rowIndex]);
                    newStyle[columnName] = style[styleIndex];
                    // The style text of one cell
                    oldStyle[columnName] = worksheet.getStyle(columnName) as string;
                    worksheet.records[rowIndex][colIndex].element.setAttribute('style', style[styleIndex]);
                    styleIndex++;
                }
                i++;
                if (row[i] != null) {
                    if (colIndex >= worksheet.headers.length - 1) {
                        // If the pasted column is out of range, create it if possible
                        if (worksheet.options.allowInsertColumn !== false) {
                            worksheet.insertColumn();
                            // Otherwise skip the pasted data that overflows
                        } else {
                            break;
                        }
                    }
                    colIndex = worksheet.selection.rightGet(colIndex, rowIndex);
                }
            }

            j++;
            if (rows[j]) {
                if (rowIndex >= worksheet.rows.length - 1) {
                    // If the pasted row is out of range, create it if possible
                    if (worksheet.options.allowInsertRow !== false) {
                        worksheet.insertRow();
                        // Otherwise skip the pasted data that overflows
                    } else {
                        break;
                    }
                }
                rowIndex = worksheet.selection.downGet(x, rowIndex);
            }
        }

        return { records, newStyle, oldStyle, lastColumn: colIndex, lastRow: rowIndex };
    }

    /**
     * Paste text in the cells from a cell.
     * @param x - The column of the first cell.
     * @param y - The row of the first cell.
     * @param data - The text. When it is the text that the worksheet copied, the cells also get the styles of the copied cells.
     */
    paste(x: number, y: number, data: string) {
        const worksheet = this.worksheet;

        // Controls
        const isCopy = worksheet.clipboard.matches(data);
        let style = isCopy ? worksheet.clipboard.style : null;

        // Depending on the behavior
        if (isCopy) {
            data = worksheet.clipboard.data;
        }

        // Split new line. The rows and columns of the text, which can grow below to fill the selection
        const text = parseCSV(data, '\t');

        // The end of the selection, to repeat the data over it. Without a selection, the data is pasted once.
        const ex = worksheet.selection.range ? worksheet.selection.range[2] : x;
        const ey = worksheet.selection.range ? worksheet.selection.range[3] : y;

        const w = ex - x + 1;
        const h = ey - y + 1;

        // Modify data to allow wor extending paste range in multiples of input range
        const repeated = repeatToSelection(text, style, w, h);
        style = repeated.style;

        // The rows to paste, which onbeforepaste can replace
        let rows: jspreadsheet.CellValue[][] = repeated.rows;

        // Paste filter
        const ret = worksheet.dispatch(
            'onbeforepaste',
            worksheet,
            repeated.rows.map((row) => {
                return row.map((item) => {
                    return { value: item };
                });
            }),
            x,
            y
        );

        if (ret === false) {
            return false;
        } else if (typeof ret === 'string') {
            // The text to paste instead
            rows = parseCSV(ret, '\t');
        } else if (Array.isArray(ret)) {
            // The rows to paste instead, of { value } objects as the event got them, or of values
            rows = ret.map((row) => {
                return row.map((item) => {
                    return item !== null && typeof item === 'object' ? item.value : item;
                });
            });
        }

        if (x != null && y != null && rows) {
            // Index
            const colIndex = parseInt(x);
            const rowIndex = parseInt(y);

            this.insertSpaceForPaste(colIndex, rowIndex, rows);

            // Go through the columns to get the data
            const { records, newStyle, oldStyle, lastColumn, lastRow } = this.pasteRows(x, y, rows, style);

            // Select the new cells
            worksheet.selection.updateFromCoords(x, y, lastColumn, lastRow);

            // Update history
            worksheet.history.add({
                action: 'setValue',
                records,
                selection: worksheet.selection.range && [...worksheet.selection.range],
                newStyle,
                oldStyle,
            });

            // Update table
            worksheet.data.updateTable();

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

            worksheet.dispatch('onpaste', worksheet, eventRecords);

            // On after changes, when a cell changed
            if (records.length) {
                const onafterchangesRecords = records.map((record) => {
                    return {
                        x: record.x,
                        y: record.y,
                        value: record.value,
                        oldValue: record.oldValue,
                    };
                });

                worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords);
            }
        }

        worksheet.clipboard.clearBorder();
    }

    /**
     * Download CSV table
     * @param includeHeaders - Put the headers in the first row.
     * @param processed - Download the shown values instead of the data.
     */
    download(includeHeaders?: boolean, processed?: boolean) {
        const worksheet = this.worksheet;

        if (worksheet.parent.config.allowExport === false) {
            console.error('Export not allowed');
        } else {
            // Data
            let data = '';

            // Get data
            data += worksheet.clipboard.copy(false, worksheet.options.csvDelimiter, true, includeHeaders, true, undefined, processed);

            // A link that downloads the file
            const blob = new Blob([`\uFEFF${data}`], { type: 'text/csv;charset=utf-8;' });

            const pom = document.createElement('a');
            pom.setAttribute('target', '_top');
            const url = URL.createObjectURL(blob);
            pom.href = url;
            pom.setAttribute('download', `${worksheet.options.csvFileName || worksheet.options.worksheetName}.csv`);
            document.body.appendChild(pom);
            pom.click();
            pom.parentNode!.removeChild(pom);
        }
    }
}

/** A 32-bit hash of a text. */
function hashText(str: string) {
    let hash = 0,
        i,
        chr;

    if (!str || str.length === 0) {
        return hash;
    } else {
        for (i = 0; i < str.length; i++) {
            chr = str.charCodeAt(i);
            hash = (hash << 5) - hash + chr;
            hash |= 0;
        }
    }
    return hash;
}
