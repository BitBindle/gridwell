import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// jsdom has no layout, so the view has a height of 1 pixel, and the window holds the first rows and the buffer
function build(
    source: Partial<Jspreadsheet.RowSource> = {},
    options: Partial<Jspreadsheet.WorksheetOptions> = {},
    spreadsheet: Partial<Jspreadsheet.SpreadsheetOptions> = {}
) {
    const calls: number[] = [];
    const [worksheet] = jspreadsheet(root, {
        worksheets: [
            {
                minDimensions: [3, 1],
                tableOverflow: true,
                rowSource: {
                    count: 100000,
                    buffer: 20,
                    row: (y) => {
                        calls.push(y);
                        return [`a${y}`, `b${y}`, `c${y}`];
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

function attached(worksheet: Worksheet) {
    return Array.from(worksheet.tbody.querySelectorAll('tr[data-y]')).map((row) => Number(row.getAttribute('data-y')));
}

describe('The row source mode', () => {
    it('makes records only for the rows of the window, and gives the table the row count', () => {
        const { worksheet, calls } = build();

        expect(attached(worksheet)).to.eql(Array.from({ length: 21 }, (row, y) => y));
        expect(calls.length).to.equal(21);
        expect(worksheet.rows.length).to.equal(100000);
        expect(worksheet.records[50000]).to.equal(undefined);
        expect(worksheet.getValueFromCoords(1, 3)).to.equal('b3');
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('100001');
        expect(worksheet.rows[3].element.getAttribute('aria-rowindex')).to.equal('5');
        expect(worksheet.rowWindow()).to.eql([0, 20]);
    });

    it('moves the window to a row, and gives the new rows to onrowwindow and onrowsattached', () => {
        const windows: number[][] = [];
        const attachedRanges: [number, number][][] = [];
        const { worksheet } = build(
            {},
            {},
            {
                onrowwindow: (instance, from, to) => windows.push([from, to]),
                onrowsattached: (instance, ranges) => attachedRanges.push(ranges),
            }
        );

        worksheet.scrollToRow(50000);

        expect(worksheet.rowWindow()![0]).to.be.at.most(50000);
        expect(worksheet.rowWindow()![1]).to.be.at.least(50000);
        expect(worksheet.getValueFromCoords(0, 50000)).to.equal('a50000');
        expect(worksheet.records[0]).to.equal(undefined);
        expect(windows[windows.length - 1]).to.eql(worksheet.rowWindow());
        expect(attachedRanges[attachedRanges.length - 1]).to.eql([worksheet.rowWindow()]);
    });

    it('keeps the frozen rows and merges the part of a range in the window', () => {
        const { worksheet } = build({}, { freezeRows: 2, mergeCells: { B5: [2, 3] } });

        expect(attached(worksheet).slice(0, 3)).to.eql([0, 1, 2]);
        expect(worksheet.getCell('B5').getAttribute('colspan')).to.equal('2');
        expect(worksheet.getCell('B5').getAttribute('rowspan')).to.equal('3');
        expect((worksheet.rows[5].element.children[3] as HTMLElement).style.display).to.equal('none');

        worksheet.scrollToRow(90000);
        expect(attached(worksheet).slice(0, 2)).to.eql([0, 1]);
        expect(worksheet.records[4]).to.equal(undefined);
        expect(worksheet.options.mergeCells).to.eql({});
        expect(worksheet.getMerge()).to.eql({ B5: [2, 3] });
    });

    it('moves the selection with the keys over hidden rows and merged ranges, also outside the window', () => {
        const { worksheet } = build({ runs: [{ first: 3, last: 9999, hidden: true }] }, { mergeCells: { A10001: [1, 3] } });
        worksheet.updateSelectionFromCoords(0, 1, 0, 1);

        worksheet.down();
        expect(worksheet.selection.range).to.eql([0, 2, 0, 2]);
        worksheet.down();
        expect(worksheet.selection.range).to.eql([0, 10000, 0, 10000]);
        worksheet.down();
        expect(worksheet.selection.range).to.eql([0, 10003, 0, 10003]);
        worksheet.up();
        expect(worksheet.selection.range).to.eql([0, 10000, 0, 10000]);

        worksheet.down(false, true);
        expect(worksheet.selection.range).to.eql([0, 99999, 0, 99999]);
        expect(worksheet.records[99999][0].element.classList.contains('highlight-selected')).to.equal(true);
        expect(worksheet.selection.activeCell).to.equal(worksheet.records[99999][0].element);
    });

    it('grows a selection over the merged ranges of the source', () => {
        const { worksheet } = build({}, { mergeCells: { B50001: [2, 2] } });

        worksheet.updateSelectionFromCoords(0, 50000, 1, 50000);

        expect(worksheet.selection.bounds).to.eql([0, 50000, 2, 50001]);
    });

    it('keeps the heights and the hidden rows in its layout', () => {
        const { worksheet } = build({ defaultHeight: 24 });

        worksheet.setHeight(2, 50);
        expect(worksheet.getHeight(2)).to.equal('50px');
        expect(worksheet.rows[2].element.style.height).to.equal('50px');
        expect(worksheet.getHeight(70000)).to.equal('24px');
        worksheet.undo();
        expect(worksheet.getHeight(2)).to.equal('24px');

        worksheet.hideRow(4);
        expect(worksheet.rows[4].element.style.display).to.equal('none');
        expect(worksheet.rowOffset(5)).to.equal(4 * 24);
        worksheet.showRow(4);
        expect(worksheet.rowOffset(5)).to.equal(5 * 24);
    });

    it('asks the height hook, and fits the rows again', () => {
        const heights: Record<number, number> = { 1: 40 };
        const { worksheet } = build({ defaultHeight: 20, height: (y) => heights[y] || 20 });

        expect(worksheet.rows[1].element.style.height).to.equal('40px');
        expect(worksheet.rowOffset(2)).to.equal(60);

        heights[1] = 70;
        worksheet.fitRows([1]);
        expect(worksheet.rowOffset(2)).to.equal(90);
    });

    it('makes rows again from the source with refreshRows and resetRows', () => {
        let prefix = 'a';
        const { worksheet } = build({ row: (y) => [`${prefix}${y}`, '', ''] });

        prefix = 'n';
        worksheet.refreshRows(2, 3);
        expect([worksheet.getValueFromCoords(0, 1), worksheet.getValueFromCoords(0, 2), worksheet.getValueFromCoords(0, 4)]).to.eql(['a1', 'n2', 'a4']);

        worksheet.updateSelectionFromCoords(0, 3, 0, 3);
        worksheet.resetRows({ count: 10, merges: { A1: [2, 2] } });
        expect(worksheet.rows.length).to.equal(10);
        expect(worksheet.getValueFromCoords(0, 4)).to.equal('n4');
        expect(worksheet.getCell('A1').getAttribute('rowspan')).to.equal('2');
        expect(worksheet.selection.range).to.eql([0, 3, 0, 3]);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('11');
    });

    it('does nothing for an insert, a delete, a move or a sort of rows or columns', () => {
        const { worksheet } = build();

        expect(worksheet.insertRow()).to.equal(false);
        expect(worksheet.deleteRow(0)).to.equal(false);
        expect(worksheet.insertColumn()).to.equal(false);
        expect(worksheet.deleteColumn(0)).to.equal(false);
        expect(worksheet.orderBy(0, 1)).to.equal(false);
        worksheet.moveRow(0, 2);
        worksheet.moveColumn(0, 2);

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('a0');
        expect(worksheet.rows.length).to.equal(100000);
    });

    it('gives oncopy an empty text and the selected range, and copies the text of the handler', () => {
        const copies: unknown[][] = [];
        const { worksheet } = build(
            {},
            {},
            {
                oncopy: (instance, range, text) => {
                    copies.push([range, text]);
                    return 'from the application';
                },
            }
        );
        worksheet.updateSelectionFromCoords(0, 2, 1, 60000);
        // jsdom has no copy command
        const execCommand = document.execCommand;
        document.execCommand = () => true;
        try {
            worksheet.copy();
        } finally {
            document.execCommand = execCommand;
        }

        expect(copies).to.eql([[[0, 2, 1, 60000], '']]);
        expect(worksheet.clipboard.textarea.value).to.equal('from the application');
    });

    it('writes the value of a row in the window, and ignores a row outside it', () => {
        const { worksheet } = build();

        worksheet.setValueFromCoords(0, 3, 'new');
        worksheet.setValueFromCoords(0, 60000, 'ignored');

        expect(worksheet.getValueFromCoords(0, 3)).to.equal('new');
        expect(worksheet.records[60000]).to.equal(undefined);
    });

    it('detaches an open editor of a row that leaves the window', () => {
        const { worksheet } = build();
        const cell = worksheet.getCell('A2');
        worksheet.openEditor(cell);
        (cell.children[0] as HTMLTextAreaElement).value = 'typed';

        worksheet.scrollToRow(80000);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.records[1]).to.equal(undefined);
    });

    it('gives the frozen columns of new rows their place after a horizontal scroll', () => {
        const { worksheet } = build({}, { freezeColumns: 1, columns: [{ width: 100 }] });
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 400, writable: true });
        worksheet.columns.updateFreezePosition();

        worksheet.scrollToRow(70000);

        const cell = worksheet.records[70000][0].element;
        expect(cell.classList.contains('jss_freezed')).to.equal(true);
        expect(cell.style.left).to.equal(worksheet.records[worksheet.rowWindow()![0]][0].element.style.left);
        expect(cell.style.left).to.not.equal('');
    });

    it('gives the rows of overlapping runs their own heights, and splits a run for one row', () => {
        const { worksheet } = build({
            defaultHeight: 20,
            runs: [
                { first: 5, last: 10, height: 40 },
                { first: 8, last: 12, hidden: true },
                { first: 100, last: 50, height: 30 },
            ],
        });

        expect(worksheet.rowOffset(5)).to.equal(100);
        expect(worksheet.rowOffset(8)).to.equal(100 + 3 * 40);
        expect(worksheet.rowOffset(13)).to.equal(100 + 3 * 40);

        worksheet.showRow(10);
        expect(worksheet.rowOffset(11)).to.equal(100 + 3 * 40 + 20);
        worksheet.setHeight(6, 25);
        expect(worksheet.rowOffset(7)).to.equal(100 + 40 + 25);
        expect(worksheet.getHeight()).to.include.members(['40px', '25px']);
    });

    it('scales the scroll position above the largest scroll height', () => {
        const { worksheet } = build({ defaultHeight: 20, maxScrollHeight: 10000 });

        worksheet.scrollToRow(99999);

        expect(worksheet.content.scrollTop).to.be.at.most(10000);
        expect(worksheet.rowWindow()![1]).to.equal(99999);
        expect(worksheet.getValueFromCoords(0, 99999)).to.equal('a99999');
    });

    it('gives a far scroll a small window first, and all rows of the buffer after a short time', async () => {
        const { worksheet } = build({ buffer: 100, defaultHeight: 20 });
        Object.defineProperty(worksheet.content, 'scrollTop', { value: 20 * 50000, writable: true });

        worksheet.content.dispatchEvent(new window.Event('scroll'));
        const [from, to] = worksheet.rowWindow()!;
        expect(to - from).to.be.below(100);

        // The window fills 100 ms after the last far scroll
        await new Promise((resolve) => setTimeout(resolve, 150));
        const [fullFrom, fullTo] = worksheet.rowWindow()!;
        expect(fullTo - fullFrom).to.be.at.least(200);
    });

    it('renders the window when the view changes its size, and stops the observer with destroy', async () => {
        const observers: { callback: () => void; disconnected: boolean }[] = [];
        const original = (window as unknown as { ResizeObserver?: unknown }).ResizeObserver;
        const Observer = class {
            callback: () => void;
            disconnected = false;
            constructor(callback: () => void) {
                this.callback = callback;
                observers.push(this);
            }
            observe() {}
            disconnect() {
                this.disconnected = true;
            }
        };
        (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = Observer;
        try {
            const windows: number[][] = [];
            const { worksheet } = build({}, {}, { onrowwindow: (instance, from, to) => windows.push([from, to]) });
            const observer = observers.find((item) => item.callback.toString().includes('render'))!;
            Object.defineProperty(worksheet.content, 'scrollTop', { value: 20 * 30000, writable: true });

            const before = windows.length;
            observer.callback();
            expect(windows.length).to.equal(before + 1);

            // The spreadsheet is in its element after its build
            await new Promise((resolve) => setTimeout(resolve));
            jspreadsheet.destroy(root as Jspreadsheet.JspreadsheetInstanceElement);
            expect(observer.disconnected).to.equal(true);
        } finally {
            (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = original;
        }
    });

    it('keeps the focus in the grid when the row of the focus leaves the window, and gives it back', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(0, 2, 0, 2);
        worksheet.getCell('A3').focus();

        worksheet.scrollToRow(60000);
        expect(document.activeElement).to.equal(worksheet.table);

        worksheet.scrollToRow(2);
        expect(document.activeElement).to.equal(worksheet.getCell('A3'));
    });

    it('starts the window at the first row of a merged range above it, and attaches the hidden rows of a range', () => {
        const { worksheet } = build({ runs: [{ first: 60011, last: 60012, hidden: true }] }, { mergeCells: { A59900: [1, 200], B60010: [1, 5] } });

        worksheet.scrollToRow(60005);

        // The window would start at row 59985, and the range starts at row 59899
        expect(worksheet.rowWindow()![0]).to.equal(59899);
        expect(Number(worksheet.getCell('A59900').getAttribute('rowspan'))).to.equal(worksheet.rowWindow()![1] - 59899 + 1);
        expect(worksheet.rows[60010].element).to.not.equal(undefined);
        expect(worksheet.rows[60011].element.style.display).to.equal('none');
        expect(worksheet.getCell('B60010').getAttribute('rowspan')).to.equal('5');
    });

    it('replaces the merged ranges of the source with replaceMerges', () => {
        const { worksheet } = build({ runs: [{ first: 4, last: 4, hidden: true }] }, { mergeCells: { A2: [2, 2] } });

        expect(worksheet.replaceMerges({ B3: [1, 3] })).to.eql([]);

        expect(worksheet.getCell('A2').getAttribute('colspan')).to.equal(null);
        expect(worksheet.getCell('B3').getAttribute('rowspan')).to.equal('3');
        expect(worksheet.rows[4].element.style.display).to.equal('none');
    });

    it('turns off the options that read or change all rows', () => {
        const { worksheet } = build({}, { minSpareRows: 2, minSpareCols: 1, footers: [['=SUM(A1:A3)', '', '']], columnDrag: true, rowDrag: true });

        worksheet.setValueFromCoords(0, 3, 'x');

        expect(worksheet.options.minSpareRows).to.equal(0);
        expect(worksheet.options.selectionCopy).to.equal(false);
        expect(worksheet.options.columnDrag).to.equal(false);
        expect(worksheet.options.allowInsertRow).to.equal(false);
        expect(worksheet.tfoot ?? null).to.equal(null);
        expect(worksheet.headers.length).to.equal(3);
    });

    it('calculates a formula of the window, and drops the formula cells of rows that leave it', () => {
        const { worksheet } = build({ row: (y) => (y === 10 ? ['', '=A1*2', ''] : [y === 0 ? 5 : '', '', '']) }, { freezeRows: 1 });
        expect(worksheet.getValueFromCoords(1, 10, true)).to.equal('10');

        worksheet.scrollToRow(80000);
        expect(Object.keys(worksheet.formulas.chain)).to.eql([]);
        worksheet.setValueFromCoords(0, 0, 7);

        worksheet.scrollToRow(10);
        expect(worksheet.getValueFromCoords(1, 10, true)).to.equal('14');
    });

    it('pastes up to the last row, and gives the values of the window to a copy with returnData', () => {
        const { worksheet } = build({ count: 100 });
        worksheet.scrollToRow(99);

        worksheet.paste(0, 98, 'p\nq\nr');
        expect([worksheet.getValueFromCoords(0, 98), worksheet.getValueFromCoords(0, 99)]).to.eql(['p', 'q']);
        expect(worksheet.rows.length).to.equal(100);

        worksheet.updateSelectionFromCoords(0, 98, 1, 99);
        expect(worksheet.clipboard.copy(true, ',', true)).to.equal('p,b98\r\nq,b99\r\n');
        worksheet.download();
        worksheet.pointer.fillDown();
        expect(worksheet.getSelected().length).to.equal(4);
        worksheet.updateSelectionFromCoords(0, 2, 1, 60);
        expect(worksheet.getSelected().every(Boolean)).to.equal(true);
    });

    it('changes the frozen rows with setFreeze', () => {
        const { worksheet } = build();
        worksheet.scrollToRow(80000);

        worksheet.setFreeze(3);

        expect(attached(worksheet).slice(0, 3)).to.eql([0, 1, 2]);
        expect(worksheet.rows[2].element.classList.contains('jss_frozen_row')).to.equal(true);
        expect(worksheet.records[80000]).to.not.equal(undefined);
    });

    it('keeps the merged ranges of setMerge, removeMerge and destroyMerge when the window moves', () => {
        const { worksheet } = build({}, { mergeCells: { B5: [2, 2] } });

        worksheet.setMerge('A10', 2, 2);
        worksheet.removeMerge('B5');
        worksheet.scrollToRow(80000);
        worksheet.scrollToRow(0);

        expect(worksheet.getMerge()).to.eql({ A10: [2, 2] });
        expect(worksheet.getCell('B5').getAttribute('colspan')).to.equal(null);

        worksheet.destroyMerge();
        worksheet.scrollToRow(80000);
        worksheet.scrollToRow(0);
        expect(worksheet.getMerge()).to.eql({});
    });

    it('has one tab stop: the active cell, or the table while the active cell is outside the window', () => {
        const { worksheet } = build();
        const stops = () => Array.from(root.querySelectorAll('[tabindex="0"]')).filter((element) => worksheet.element.contains(element));

        worksheet.updateSelectionFromCoords(0, 3, 0, 3);
        expect(stops()).to.eql([worksheet.getCell('A4')]);

        worksheet.scrollToRow(80000);
        expect(stops()).to.eql([worksheet.table]);
    });

    it('gives the focus of an open editor to the table when its row leaves the window', () => {
        const { worksheet } = build();
        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        const cell = worksheet.getCell('A2');
        worksheet.openEditor(cell);
        (cell.children[0] as HTMLTextAreaElement).focus();

        worksheet.scrollToRow(80000);
        expect(document.activeElement).to.equal(worksheet.table);

        worksheet.scrollToRow(1);
        expect(document.activeElement).to.equal(worksheet.getCell('A2'));
    });

    it('lets a run that starts inside an earlier run win for its rows', () => {
        const { worksheet } = build({
            count: 10,
            defaultHeight: 20,
            runs: [
                { first: 0, last: 4, height: 30 },
                { first: 0, last: 1, height: 40 },
                { first: 0, last: 1, height: 50 },
            ],
        });

        expect(worksheet.rowOffset(2)).to.equal(100);
        expect(worksheet.rowOffset(5)).to.equal(190);
        expect(worksheet.rowOffset(10)).to.equal(290);
    });

    it('keeps the scroll area below the limit of sticky positioning in Firefox', () => {
        const { worksheet } = build({ count: 1048576, defaultHeight: 20, maxScrollHeight: 50000000 });

        worksheet.scrollToRow(1048575);

        expect(worksheet.content.scrollTop).to.be.below(8000000);
    });

    it('resizes a column with the mouse', async () => {
        const { worksheet } = build({}, { columns: [{ width: 100 }, { width: 100 }, { width: 100 }] });
        // The spreadsheet adds its event handlers after its build
        await new Promise((resolve) => setTimeout(resolve));
        const header = worksheet.headers[1];
        header.getBoundingClientRect = () => ({ left: 0, top: 0, right: 100, bottom: 20, width: 100, height: 20 }) as DOMRect;
        function mouse(type: string, offsetX: number, pageX: number) {
            const event = new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1 });
            Object.defineProperty(event, 'offsetX', { value: offsetX });
            Object.defineProperty(event, 'pageX', { value: pageX });
            header.dispatchEvent(event);
        }

        mouse('mousedown', 98, 98);
        mouse('mousemove', 0, 148);
        mouse('mouseup', 0, 148);

        expect(worksheet.getWidth(1)).to.equal(150);
        expect(worksheet.rows.length).to.equal(100000);
    });

    it('sets a radio cell of the window', () => {
        const { worksheet } = build({ row: () => [false, '', ''] }, { columns: [{ type: 'radio' }, {}, {}] });

        worksheet.setValueFromCoords(0, 3, true);

        expect(worksheet.getValueFromCoords(0, 3)).to.equal(true);
    });

    it('merges the selection with setMerge, and refuses a range of one cell or a range that crosses another range', () => {
        const alerts: string[] = [];
        const { worksheet } = build({}, { mergeCells: { A2: [2, 2] } }, { alert: (message) => alerts.push(message) });

        worksheet.setMerge('B3', 2, 2);
        worksheet.setMerge('C10', 1, 1);
        expect(worksheet.getMerge()).to.eql({ A2: [2, 2] });
        expect(alerts).to.eql(['There is a conflict with another merged cell', 'Invalid merged properties']);

        worksheet.updateSelectionFromCoords(0, 5, 1, 5000);
        worksheet.setMerge();
        expect(worksheet.getMerge('A6')).to.eql([2, 4996]);

        // The merged ranges outside the window stay in getMerge
        worksheet.scrollToRow(80000);
        expect(worksheet.getMerge()).to.eql({ A2: [2, 2], A6: [2, 4996] });
    });

    it('removes many merged ranges at once with destroyMerge', () => {
        const mergeCells: Record<string, [number, number]> = {};
        for (let k = 0; k < 2000; k++) {
            mergeCells[`A${k * 3 + 1}`] = [2, 2];
        }
        const events: string[] = [];
        const { worksheet } = build({}, { mergeCells }, { onunmerge: (instance, cellName) => events.push(cellName) });

        worksheet.destroyMerge();

        expect(worksheet.getMerge()).to.eql({});
        expect(events.length).to.equal(2000);
    });

    it('gives many single rows inside one run their own state', () => {
        const runs: Jspreadsheet.RowRun[] = [{ first: 0, last: 99999, height: 30 }];
        for (let k = 1; k <= 5000; k++) {
            runs.push({ first: k * 10, last: k * 10, hidden: true });
        }
        const { worksheet } = build({ defaultHeight: 20, runs });

        expect(worksheet.rowOffset(100000)).to.equal((100000 - 5000) * 30);
        expect(worksheet.rowOffset(11)).to.equal(10 * 30);
    });

    it('shows no Save as item in the context menu', async () => {
        let titles: (string | undefined)[] = [];
        const { worksheet } = build({}, {}, { contextMenu: (instance, x, y, e, items) => void (titles = items.map((item) => item.title)) });
        await new Promise((resolve) => setTimeout(resolve));

        worksheet.getCell('A2').dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(titles.length).to.be.above(0);
        expect(titles).to.not.include('Save as...');
    });

    it('refuses with replaceMerges and setMerge the merged ranges that the other modes refuse', () => {
        const alerts: string[] = [];
        const { worksheet } = build({ count: 50 }, { mergeCells: { A2: [2, 2], B3: [2, 2] } }, { alert: (message) => alerts.push(message) });
        expect(worksheet.getMerge()).to.eql({ A2: [2, 2] });

        const refused = worksheet.replaceMerges({
            foo: [2, 2],
            A1: [1, 1],
            B2: [2, 2],
            C2: [2, 2],
            // @ts-expect-error -- spans that are not whole numbers
            A5: [2.5, 'x'],
            C8: [5, 1],
            A49: [1, 5],
        });
        expect(refused.sort()).to.eql(['A1', 'A49', 'A5', 'C2', 'C8', 'foo']);
        expect(worksheet.getMerge()).to.eql({ B2: [2, 2] });

        worksheet.setMerge('a1', 2, 2);
        worksheet.setMerge('C5', 3, 1);
        worksheet.setMerge('A48', 1, 10);
        expect(alerts).to.eql(['Invalid merged properties', 'Invalid merged properties', 'Invalid merged properties']);
        expect(worksheet.getMerge()).to.eql({ B2: [2, 2] });

        // A window move after the refused ranges works
        worksheet.scrollToRow(40);
        worksheet.scrollToRow(0);
        expect(worksheet.getCell('B2').getAttribute('rowspan')).to.equal('2');
    });

    it('leaves out the merged ranges that do not fit the new row count of resetRows', () => {
        const { worksheet } = build({ count: 50 }, { mergeCells: { A2: [2, 2], A9: [2, 5], A40: [2, 2] } });

        worksheet.resetRows({ count: 10 });

        expect(worksheet.getMerge()).to.eql({ A2: [2, 2] });
        worksheet.updateSelectionFromCoords(0, 9, 0, 9);
        expect(worksheet.selection.bounds).to.eql([0, 9, 0, 9]);
    });

    it('keeps the order of the keys and the unchanged ranges in replaceMerges, with the events of the other modes', () => {
        const events: string[] = [];
        const { worksheet } = build(
            {},
            { mergeCells: { A2: [2, 2], C5: [1, 2] } },
            {
                onmerge: (instance, merges) => events.push(`merge ${Object.keys(merges).join(',')}`),
                onunmerge: (instance, cellName) => events.push(`unmerge ${cellName}`),
            }
        );

        // A2 does not change, so A1 crosses it. B5 is before C6 in the order of the keys.
        expect(worksheet.replaceMerges({ A1: [1, 2], A2: [2, 2], B5: [2, 2], C6: [1, 2] }).sort()).to.eql(['A1', 'C6']);
        expect(worksheet.getMerge()).to.eql({ A2: [2, 2], B5: [2, 2] });
        expect(events).to.eql(['unmerge C5', 'merge B5']);
    });

    it('merges a span below 1 as one cell in setMerge, and refuses a cell that a range has', () => {
        const alerts: string[] = [];
        const { worksheet } = build({}, {}, { alert: (message) => alerts.push(message) });

        worksheet.setMerge('A1', 0, 2);
        worksheet.setMerge('A10', 2, -1);
        expect(worksheet.getMerge()).to.eql({ A1: [1, 2], A10: [2, 1] });

        worksheet.setMerge('A1', 2, 2);
        expect(alerts).to.eql(['Cell already merged']);
    });

    it('moves the selection, the tab stop and an open editor out of the cells that a new merged range hides', () => {
        const { worksheet } = build();
        const stops = () => Array.from(root.querySelectorAll('[tabindex="0"]')).filter((element) => worksheet.element.contains(element));

        // setMerge selects the first cell of the range, also for a selection from the bottom right
        worksheet.updateSelectionFromCoords(1, 5, 0, 2);
        worksheet.setMerge();
        expect(worksheet.getMerge()).to.eql({ A3: [2, 4] });
        expect(worksheet.selection.range?.slice(0, 2)).to.eql([0, 2]);
        expect(stops()).to.eql([worksheet.getCell('A3')]);

        // replaceMerges selects the first cell of a range that hides the selected cell, and closes its editor without a save
        worksheet.updateSelectionFromCoords(1, 9, 1, 9);
        const cell = worksheet.records[9][1].element;
        worksheet.openEditor(cell);
        (cell.children[0] as HTMLTextAreaElement).value = 'typed';
        worksheet.replaceMerges({ A3: [2, 4], A9: [2, 2] });
        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(1, 9)).to.equal('b9');
        expect(worksheet.selection.range?.slice(0, 2)).to.eql([0, 8]);
        expect(stops()).to.eql([worksheet.getCell('A9')]);
    });

    it('moves the tab stop out of a hidden active cell that is not the selected cell', () => {
        const { worksheet } = build();
        const stops = () => Array.from(root.querySelectorAll('[tabindex="0"]')).filter((element) => worksheet.element.contains(element));
        const visible = () => stops().every((element) => (element as HTMLElement).style.display !== 'none');

        // No cell is selected, and the active cell stays
        worksheet.updateSelectionFromCoords(1, 2, 1, 2);
        worksheet.resetSelection();
        worksheet.replaceMerges({ A2: [2, 2] });
        expect(stops()).to.have.length(1);
        expect(visible()).to.equal(true);

        // The selected row is outside the window
        worksheet.replaceMerges({});
        worksheet.updateSelectionFromCoords(1, 5, 1, 5);
        worksheet.updateSelectionFromCoords(1, 499, 1, 499);
        worksheet.replaceMerges({ A5: [2, 2] });
        expect(visible()).to.equal(true);
    });

    it('moves the selection before onunmerge, as the other modes do', () => {
        const events: string[] = [];
        const { worksheet } = build(
            {},
            {},
            {
                onselection: (_, x, y) => events.push(`selection ${x},${y}`),
                onunmerge: (_, name) => events.push(`unmerge ${name}`),
                onmerge: (_, merges) => events.push(`merge ${Object.keys(merges).join()}`),
            }
        );
        worksheet.replaceMerges({ A2: [2, 2] });
        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        events.length = 0;
        worksheet.replaceMerges({ A3: [2, 2] });
        expect(events).to.eql(['selection 0,1', 'unmerge A2', 'merge A3']);
    });

    it('checks merged ranges whose last row is above 2^31', () => {
        const { worksheet } = build({ count: 3000000000 });
        const refused = worksheet.replaceMerges({ A1: [1, 2147483650], A100: [2, 2] });
        expect(refused.length).to.equal(1);
    });
});
