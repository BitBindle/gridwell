import type { Spreadsheet } from '../types/internal';

/**
 * Get carret position for one element
 */
export function getCaretIndex(this: Spreadsheet, e: Node) {
    // Applications call it as jspreadsheet.helpers.getCaretIndex, so this is a spreadsheet only when the caller sets it.
    // The root option can be a shadow root, which has getSelection only in some browsers, or an element, which never has
    // it. Without it, the selection of the window has the caret.
    const root = this?.config?.root as (Node & { getSelection?: () => Selection | null }) | undefined;
    const s = typeof root?.getSelection === 'function' ? root.getSelection() : window.getSelection();

    let pos = 0;
    if (s) {
        if (s.rangeCount !== 0) {
            const r = s.getRangeAt(0);
            const p = r.cloneRange();
            p.selectNodeContents(e);
            p.setEnd(r.endContainer, r.endOffset);
            pos = p.toString().length;
        }
    }
    return pos;
}

/**
 * Invert keys and values
 */
export function invert(o: Record<string, string | number>) {
    // The values become the indexes, or the keys when they are not numbers
    const d = [] as unknown as string[] & Record<string, string>;
    for (const [key, value] of Object.entries(o)) {
        d[value] = key;
    }
    return d;
}

/**
 * Get the letters of a column, for example "B" for 1.
 * @param columnNumber - The column, from 0.
 */
export function getColumnName(columnNumber: number) {
    let dividend = columnNumber + 1;
    let columnName = '';
    let modulo;

    while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo).toString() + columnName;
        dividend = parseInt((dividend - modulo) / 26);
    }

    return columnName;
}

/**
 * Get column name from coords
 */
export function getCellNameFromCoords(x: number, y: number) {
    return getColumnName(parseInt(x)) + (parseInt(y) + 1);
}

export function getCoordsFromCellName(columnName: string): [number, number | null] | undefined {
    // Get the letters
    const t = /^[a-zA-Z]+/.exec(columnName);

    if (t) {
        // Base 26 calculation
        let code = 0;
        for (let i = 0; i < t[0].length; i++) {
            code += parseInt(t[0].charCodeAt(i) - 64) * Math.pow(26, t[0].length - 1 - i);
        }
        code--;
        // Make sure jspreadsheet starts on zero
        if (code < 0) {
            code = 0;
        }

        // Number
        // parseInt converts the match to a string
        let number = parseInt(/[0-9]+$/.exec(columnName) as unknown as string) || null;
        if (number! > 0) {
            number!--;
        }

        return [code, number];
    }
}

export function getCoordsFromRange(range: string): [number, number, number, number] {
    const [start, end] = range.split(':');

    // A range of whole columns, for example A:B, has no rows
    return [...getCoordsFromCellName(start)!, ...getCoordsFromCellName(end)!] as [number, number, number, number];
}

/**
 * From stack overflow contributions
 */
export function parseCSV(str: string, delimiter?: string) {
    // user-supplied delimeter or default comma
    delimiter = delimiter || ',';
    // Remove last line break
    str = str.replace(/\r?\n$|\r$|\n$/g, '');

    const arr: string[][] = [];
    let quote = false; // true means we're inside a quoted field
    // iterate over each character, keep track of current row and column (of the returned array)
    let maxCol = 0;
    let row = 0,
        col = 0;
    for (let c = 0; c < str.length; c++) {
        const cc = str[c],
            nc = str[c + 1];
        arr[row] = arr[row] || [];
        arr[row][col] = arr[row][col] || '';

        // If the current character is a quotation mark, and we're inside a quoted field, and the next character is also a quotation mark, add a quotation mark to the current column and skip the next character
        if (cc === '"' && quote && nc === '"') {
            arr[row][col] += cc;
            ++c;
            continue;
        }

        // If it's just one quotation mark, begin/end quoted field
        if (cc === '"') {
            quote = !quote;
            continue;
        }

        // If it's a comma and we're not in a quoted field, move on to the next column
        if (cc === delimiter && !quote) {
            ++col;
            continue;
        }

        // If it's a newline (CRLF) and we're not in a quoted field, skip the next character and move on to the next row and move to column 0 of that new row
        if (cc === '\r' && nc === '\n' && !quote) {
            ++row;
            maxCol = Math.max(maxCol, col);
            col = 0;
            ++c;
            continue;
        }

        // If it's a newline (LF or CR) and we're not in a quoted field, move on to the next row and move to column 0 of that new row
        if (cc === '\n' && !quote) {
            ++row;
            maxCol = Math.max(maxCol, col);
            col = 0;
            continue;
        }
        if (cc === '\r' && !quote) {
            ++row;
            maxCol = Math.max(maxCol, col);
            col = 0;
            continue;
        }

        // Otherwise, append the current character to the current column
        arr[row][col] += cc;
    }

    // fix array length
    arr.forEach((row, i) => {
        for (let i = row.length; i <= maxCol; i++) {
            row.push('');
        }
    });
    return arr;
}

// The options are optional here: factory.ts passes the first worksheet of the options, which can be missing
export function createFromTable(el: HTMLElement, options?: jspreadsheet.WorksheetOptions) {
    if (el.tagName !== 'TABLE') {
        console.log('Element is not a table');
    } else {
        // Configuration
        if (!options) {
            options = {};
        }

        // The columns and the rows of the table are added to these arrays
        const columns: jspreadsheet.Column[] = [];
        const data: string[][] = [];
        options.columns = columns;
        options.data = data;

        // Colgroup
        const colgroup = el.querySelectorAll<HTMLElement>('colgroup > col');
        if (colgroup.length) {
            // Get column width
            for (let i = 0; i < colgroup.length; i++) {
                let width = colgroup[i].style.width;
                if (!width) {
                    width = colgroup[i].getAttribute('width')!;
                }
                // Set column width
                if (width) {
                    if (!columns[i]) {
                        columns[i] = {};
                    }
                    columns[i].width = width;
                }
            }
        }

        // Parse header
        function parseHeader(header: HTMLElement, i: number) {
            // Get width information
            const rect = header.getBoundingClientRect();
            const width = rect.width > 50 ? rect.width : 50;

            // Create column option
            if (!columns[i]) {
                columns[i] = {};
            }
            if (header.getAttribute('data-celltype')) {
                columns[i].type = header.getAttribute('data-celltype') as jspreadsheet.Column['type'];
            } else {
                columns[i].type = 'text';
            }
            columns[i].width = `${width}px`;
            columns[i].title = header.innerHTML;
            if (header.style.textAlign) {
                columns[i].align = header.style.textAlign as jspreadsheet.HorizontalAlign;
            }

            let info;
            if ((info = header.getAttribute('name'))) {
                columns[i].name = info;
            }
            if ((info = header.getAttribute('id'))) {
                // id is not in the typings of the columns
                (columns[i] as { id?: string }).id = info;
            }
            if ((info = header.getAttribute('data-mask'))) {
                columns[i].mask = info;
            }
        }

        // Headers
        const nested = [];
        // The rows of the thead, then the cells of its last row
        let headers: NodeListOf<Element> | HTMLCollection = el.querySelectorAll(':scope > thead > tr');
        if (headers.length) {
            for (let j = 0; j < headers.length - 1; j++) {
                const cells = [];
                for (const child of headers[j].children) {
                    const row = {
                        title: child.textContent,
                        colspan: parseInt(child.getAttribute('colspan')) || 1,
                    };
                    cells.push(row);
                }
                nested.push(cells);
            }
            // Get the last row in the thead
            headers = headers[headers.length - 1].children;
            // Go though the headers
            for (let i = 0; i < headers.length; i++) {
                parseHeader(headers[i] as HTMLElement, i);
            }
        }

        // Content
        let rowNumber = 0;
        const mergeCells: Record<string, [number, number]> = {};
        const rows: Record<number, jspreadsheet.Row> = {};
        const style: Record<string, string> = {};
        const classes: Record<string, string> = {};

        let content = el.querySelectorAll<HTMLElement>(':scope > tr, :scope > tbody > tr');
        for (let j = 0; j < content.length; j++) {
            data[rowNumber] = [];
            if (options.parseTableFirstRowAsHeader === true && !headers.length && j === 0) {
                for (let i = 0; i < content[j].children.length; i++) {
                    parseHeader(content[j].children[i] as HTMLElement, i);
                }
            } else {
                for (let i = 0; i < content[j].children.length; i++) {
                    // WickedGrid formula compatibility
                    let value = content[j].children[i].getAttribute('data-formula');
                    if (value) {
                        if (value.substr(0, 1) !== '=') {
                            value = `=${value}`;
                        }
                    } else {
                        value = content[j].children[i].innerHTML;
                    }
                    data[rowNumber].push(value);

                    // Key, by the row in the data, which is j - 1 when the first row is the header
                    const cellName = getCellNameFromCoords(i, rowNumber);

                    // Classes
                    const tmp = content[j].children[i].getAttribute('class');
                    if (tmp) {
                        classes[cellName] = tmp;
                    }

                    // Merged cells
                    const mergedColspan = parseInt(content[j].children[i].getAttribute('colspan')) || 0;
                    const mergedRowspan = parseInt(content[j].children[i].getAttribute('rowspan')) || 0;
                    if (mergedColspan || mergedRowspan) {
                        mergeCells[cellName] = [mergedColspan || 1, mergedRowspan || 1];
                    }

                    // Avoid problems with hidden cells
                    if ((content[j].children[i] as HTMLElement).style?.display === 'none') {
                        (content[j].children[i] as HTMLElement).style.display = '';
                    }
                    // Get style
                    const s = content[j].children[i].getAttribute('style');
                    if (s) {
                        style[cellName] = s;
                    }
                    // Bold
                    if (content[j].children[i].classList.contains('styleBold')) {
                        if (style[cellName]) {
                            style[cellName] += '; font-weight:bold;';
                        } else {
                            style[cellName] = 'font-weight:bold;';
                        }
                    }
                }

                // Row Height
                if (content[j].style?.height) {
                    rows[rowNumber] = { height: content[j].style.height };
                }

                // Index
                rowNumber++;
            }
        }

        // Nested
        if (Object.keys(nested).length > 0) {
            options.nestedHeaders = nested;
        }
        // Style
        if (Object.keys(style).length > 0) {
            options.style = style;
        }
        // Merged
        if (Object.keys(mergeCells).length > 0) {
            options.mergeCells = mergeCells;
        }
        // Row height
        if (Object.keys(rows).length > 0) {
            options.rows = rows;
        }
        // Classes
        if (Object.keys(classes).length > 0) {
            options.classes = classes;
        }

        content = el.querySelectorAll<HTMLElement>('tfoot tr');
        if (content.length) {
            const footers = [];
            for (const rowElement of content) {
                const footer = [];
                for (const cell of rowElement.children) {
                    footer.push(cell.textContent);
                }
                footers.push(footer);
            }
            if (Object.keys(footers).length > 0) {
                options.footers = footers;
            }
        }
        // TODO: data-hiddencolumns="3,4"

        // I guess in terms the better column type
        if (options.parseTableAutoCellType === true) {
            // The number of times each value is in each column
            const pattern: Record<string, number>[] = [];
            for (let i = 0; i < columns.length; i++) {
                let test = true;
                let testCalendar = true;
                // An array used as an object, with the values as keys
                pattern[i] = [] as unknown as Record<string, number>;
                for (const dataRow of data) {
                    const value = dataRow[i];
                    if (!pattern[i][value]) {
                        pattern[i][value] = 0;
                    }
                    pattern[i][value]++;
                    if (value.length > 25) {
                        test = false;
                    }
                    if (value.length === 10) {
                        if (!(value.substr(4, 1) === '-' && value.substr(7, 1) === '-')) {
                            testCalendar = false;
                        }
                    } else {
                        testCalendar = false;
                    }
                }

                const keys = Object.keys(pattern[i]).length;
                if (testCalendar) {
                    columns[i].type = 'calendar';
                } else if (test === true && keys > 1 && keys <= parseInt(data.length * 0.1)) {
                    columns[i].type = 'dropdown';
                    (columns[i] as jspreadsheet.DropdownColumn).source = Object.keys(pattern[i]);
                }
            }
        }

        return options;
    }
}
