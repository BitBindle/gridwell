import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// jsdom has no layout, so the view has a width of 1 pixel, and the column window holds the column at the left edge of
// the view and the buffer. The worksheet has 200 columns of 100 pixels.
const columnCount = 200;

function build(
    source: Partial<Jspreadsheet.RowSource> = {},
    options: Partial<Jspreadsheet.WorksheetOptions> = {},
    spreadsheet: Partial<Jspreadsheet.SpreadsheetOptions> = {}
) {
    const calls: { y: number; columns?: number[] }[] = [];
    const [worksheet] = jspreadsheet(root, {
        worksheets: [
            {
                minDimensions: [columnCount, 1],
                tableOverflow: true,
                rowSource: {
                    count: 1000,
                    buffer: 5,
                    columnWindow: true,
                    row: (y, columns) => {
                        calls.push({ y, columns: columns && [...columns] });
                        const values: Jspreadsheet.CellValue[] = [];
                        for (const x of columns || Array.from({ length: columnCount }, (_, i) => i)) {
                            values[x] = `${x}:${y}`;
                        }
                        return values;
                    },
                    ...source,
                },
                ...options,
            },
        ],
        ...spreadsheet,
    });
    return { worksheet, calls };
}

/** The columns with a cell in row y, in the order of the table. */
function cellColumns(worksheet: Worksheet, y: number) {
    return Array.from(worksheet.rows[y].element.querySelectorAll('td[data-x]')).map((td) => Number(td.getAttribute('data-x')));
}

/** The columns with a header, in the order of the table. */
function headerColumns(worksheet: Worksheet) {
    return Array.from(worksheet.headerContainer.querySelectorAll('td[data-x]')).map((td) => Number(td.getAttribute('data-x')));
}

function range(from: number, to: number) {
    return Array.from({ length: to - from + 1 }, (_, i) => from + i);
}

function key(keyCode: number, name: string) {
    document.dispatchEvent(new KeyboardEvent('keydown', { key: name, keyCode, which: keyCode, bubbles: true }));
}

describe('The column window of the row source mode', () => {
    it('makes cells only for the columns of the window, and asks the row source only for these columns', () => {
        const { worksheet, calls } = build();

        expect(worksheet.columnWindow()).to.eql([0, 10]);
        expect(cellColumns(worksheet, 0)).to.eql(range(0, 10));
        expect(headerColumns(worksheet)).to.eql(range(0, 10));
        expect(worksheet.headers.length).to.equal(columnCount);
        expect(worksheet.records[0][50]).to.equal(undefined);
        expect(calls[0].columns).to.eql(range(0, 10));
        expect(worksheet.getValueFromCoords(3, 2)).to.equal('3:2');

        // Two spacer cells give the row the width of all columns
        const spacers = worksheet.colgroupContainer.querySelectorAll('col.jss_colspacer');
        expect(spacers.length).to.equal(2);
        expect(Number(spacers[0].getAttribute('width')) + Number(spacers[1].getAttribute('width'))).to.equal((columnCount - 11) * 100);
        expect(worksheet.rows[0].element.querySelectorAll('td.jss_colspacer').length).to.equal(2);
    });

    it('moves the window to a column, and gives the new columns to oncolumnwindow and oncolumnsattached', () => {
        const windows: number[][] = [];
        const attachedRanges: [number, number][][] = [];
        const { worksheet, calls } = build(
            {},
            {},
            {
                oncolumnwindow: (instance, from, to, first, last) => windows.push([from, to, first, last]),
                oncolumnsattached: (instance, ranges, rows) => attachedRanges.push([...ranges, ...rows]),
            }
        );
        calls.length = 0;
        windows.length = 0;
        attachedRanges.length = 0;

        worksheet.scrollToColumn(150);

        // A column wider than the view goes to the left edge of the view
        expect(worksheet.columnWindow()).to.eql([140, 160]);
        expect(windows).to.eql([[140, 160, 150, 150]]);
        expect(attachedRanges).to.eql([
            [
                [140, 160],
                [0, 5],
            ],
        ]);
        expect(cellColumns(worksheet, 3)).to.eql(range(140, 160));
        expect(headerColumns(worksheet)).to.eql(range(140, 160));
        expect(calls.map((call) => call.columns)).to.eql(range(0, 5).map(() => range(140, 160)));
        expect(worksheet.getValueFromCoords(150, 3)).to.equal('150:3');
        expect(worksheet.records[3][5]).to.equal(undefined);

        // getData gives empty values for the columns without a cell
        const row = worksheet.getData()[3];
        expect(row.length).to.equal(columnCount);
        expect(row[5]).to.equal('');
        expect(row[150]).to.equal('150:3');

        // A column that shows does not move the window
        worksheet.scrollToColumn(150);
        expect(windows.length).to.equal(1);
    });

    it('keeps the frozen columns in each row', () => {
        const { worksheet } = build({}, { freezeColumns: 2 });

        worksheet.scrollToColumn(100);

        const [from, to] = worksheet.columnWindow()!;
        expect(from).to.be.above(2);
        expect(cellColumns(worksheet, 0)).to.eql([0, 1, ...range(from, to)]);
        expect(headerColumns(worksheet)).to.eql([0, 1, ...range(from, to)]);
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('1:0');
    });

    it('keeps a change of a cell of the window, and ignores a change of a cell outside it', () => {
        const { worksheet } = build();

        worksheet.setValueFromCoords(2, 1, 'changed');
        worksheet.setValueFromCoords(100, 1, 'outside');

        expect(worksheet.getValueFromCoords(2, 1)).to.equal('changed');
        expect(worksheet.records[1][100]).to.equal(undefined);
        expect(worksheet.getData()[1][100]).to.equal('');
    });

    it('gives the selection classes to the cells that the window adds, and leaves out the other cells in getSelected', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(2, 1, 150, 2);

        expect(worksheet.getSelected().length).to.equal(2 * 9);
        expect(worksheet.getSelected(true).length).to.equal(2 * 149);

        worksheet.scrollToColumn(150);
        expect(worksheet.records[1][148].element.classList.contains('highlight')).to.equal(true);
        expect(worksheet.records[1][150].element.classList.contains('highlight-right')).to.equal(true);
        expect(worksheet.records[1][155].element.classList.contains('highlight')).to.equal(false);
        expect(worksheet.headers[148].classList.contains('selected')).to.equal(true);
    });

    it('scrolls the window with the arrow keys, and gives the tab stop to the selected cell', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(10, 1, 10, 1);

        key(39, 'ArrowRight');

        expect(worksheet.selection.range).to.eql([11, 1, 11, 1]);
        expect(worksheet.columnWindow()![1]).to.be.at.least(11);
        const cell = worksheet.records[1][11].element;
        expect(cell.classList.contains('highlight-selected')).to.equal(true);
        expect(cell.tabIndex).to.equal(0);

        key(35, 'End');
        expect(worksheet.selection.range![0]).to.equal(columnCount - 1);
        expect(worksheet.columnWindow()![1]).to.equal(columnCount - 1);
        expect(worksheet.records[1][columnCount - 1].element.tabIndex).to.equal(0);
    });

    it('starts the window at the first cell of a merged range left of it, or shows the text of a range that starts further left', () => {
        const { worksheet } = build();
        worksheet.setMerge('C3', 50, 2);

        worksheet.scrollToColumn(60);

        // The range from column 2 is 48 columns left of the window, so the window starts at its first cell
        const [from] = worksheet.columnWindow()!;
        expect(from).to.equal(2);
        expect(worksheet.records[2][2].element.getAttribute('colspan')).to.equal('50');

        worksheet.setMerge('A5', 150, 1);
        worksheet.scrollToColumn(145);

        // The range from column 0 is 135 columns left of the window: its part in the window merges, and shows its text
        const [start] = worksheet.columnWindow()!;
        expect(start).to.equal(135);
        const cell = worksheet.records[4][135].element;
        expect(cell.getAttribute('colspan')).to.equal('15');
        expect(cell.textContent).to.equal('0:4');
        expect(worksheet.getValueFromCoords(135, 4)).to.equal('135:4');

        // The cell shows its own value when the range goes
        worksheet.removeMerge('A5');
        expect(worksheet.records[4][135].element.textContent).to.equal('135:4');
    });

    it('reads the widths and the hidden columns again after setWidth, hideColumn and refreshColumns', () => {
        const { worksheet } = build();
        const spacerWidth = () => {
            const spacers = worksheet.colgroupContainer.querySelectorAll('col.jss_colspacer');
            return Number(spacers[0].getAttribute('width')) + Number(spacers[1].getAttribute('width'));
        };
        expect(spacerWidth()).to.equal(189 * 100);

        worksheet.setWidth(100, 300);
        expect(spacerWidth()).to.equal(188 * 100 + 300);

        worksheet.hideColumn(120);
        expect(spacerWidth()).to.equal(187 * 100 + 300);

        worksheet.showColumn(120);
        worksheet.cols[130].colElement.setAttribute('width', '50');
        worksheet.refreshColumns();
        expect(spacerWidth()).to.equal(187 * 100 + 300 + 50);
    });

    it('keeps the scroll area below maxScrollWidth, and maps a scroll position to a column in proportion', () => {
        const { worksheet } = build({ maxScrollWidth: 10000 });
        const spacers = worksheet.colgroupContainer.querySelectorAll('col.jss_colspacer');
        const width = () =>
            Number(spacers[0].getAttribute('width')) +
            Number(spacers[1].getAttribute('width')) +
            (worksheet.columnWindow()![1] - worksheet.columnWindow()![0] + 1) * 100;
        expect(width()).to.be.at.most(10000);

        worksheet.scrollToColumn(150);
        expect(worksheet.columnWindow()![0]).to.be.at.most(150);
        expect(worksheet.columnWindow()![1]).to.be.at.least(150);
        expect(width()).to.be.at.most(10000);
    });

    it('gives the focus of an open editor to the table when its column leaves the window', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(2, 1, 2, 1);
        worksheet.openEditor(worksheet.records[1][2].element);
        (worksheet.records[1][2].element.querySelector('textarea, input') as HTMLElement).focus();

        worksheet.scrollToColumn(150);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.records[1][2]).to.equal(undefined);
        expect(document.activeElement).to.equal(worksheet.table);
    });

    it('stops the text of a new cell at a right neighbour with a value when textOverflow is on', () => {
        const { worksheet } = build({}, { textOverflow: true });

        worksheet.scrollToColumn(150);
        // The window moves left, so the new cells come before the cells that it kept, from column 140
        worksheet.scrollToColumn(135);

        const [from] = worksheet.columnWindow()!;
        expect(from).to.equal(125);
        for (let x = from; x < 140; x++) {
            expect(worksheet.records[2][x].element.style.overflow, `column ${x}`).to.equal('hidden');
        }
    });

    it('turns off the nested headers', () => {
        const { worksheet } = build({}, { nestedHeaders: [[{ title: 'All', colspan: columnCount }]] });

        expect(worksheet.options.nestedHeaders).to.equal(undefined);
        expect(worksheet.thead.querySelectorAll('tr').length).to.equal(1);
    });

    it('reads the comments, styles and data of the window, and sets the style, comments and read-only state of its cells', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(0, 0, columnCount - 1, 3);

        worksheet.setStyle({ B2: 'color: red', EA2: 'color: blue' });
        worksheet.setComments({ B2: 'note', EA2: 'far' });
        worksheet.setReadOnly('EA2', true);
        worksheet.setReadOnly('B2', true);

        expect(worksheet.getStyle('B2')).to.contain('color: red');
        expect(worksheet.getComments('B2')).to.equal('note');
        expect(worksheet.isReadOnly(1, 1)).to.equal(true);
        expect(Object.keys(worksheet.getStyle() as Record<string, string>)).to.not.contain('EA2');
        expect(Object.keys(worksheet.getComments() as Record<string, string>)).to.eql(['B2']);
        expect(worksheet.getColumnData(150, true).every((value) => value === '')).to.equal(true);
    });

    it('gives the row source only the columns of the window after resetRows, and renders the window again', () => {
        const { worksheet, calls } = build();
        worksheet.scrollToColumn(150);
        calls.length = 0;

        worksheet.resetRows({ count: 10 });

        expect(worksheet.columnWindow()).to.eql([140, 160]);
        expect(calls.length).to.be.above(0);
        for (const call of calls) {
            expect(call.columns).to.eql(range(140, 160));
        }
        expect(cellColumns(worksheet, 0)).to.eql(range(140, 160));
    });
});
