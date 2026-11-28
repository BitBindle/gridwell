import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

const grid = (rows: number, columns: number) => {
    const data: string[][] = [];
    for (let j = 0; j < rows; j++) {
        const row: string[] = [];
        for (let i = 0; i < columns; i++) {
            row.push(`${String.fromCharCode(65 + i)}${j + 1}`);
        }
        data.push(row);
    }
    return data;
};

const load = (worksheet: Jspreadsheet.WorksheetOptions, options: Partial<Jspreadsheet.SpreadsheetOptions> = {}) => {
    const events: string[] = [];
    const instance = jspreadsheet(root, {
        ...options,
        worksheets: [worksheet],
        onevent(name) {
            events.push(name);
        },
    });
    events.length = 0;
    return { worksheet: instance[0], events };
};

// The events that also occur while ignoreEvents is set report changes of the DOM, not of the data
const dataEvents = (events: string[]) => events.filter((name) => !['onrowsattached', 'oncolumnsattached', 'oneditordetach'].includes(name));

describe('shiftRows', () => {
    let confirm: typeof global.confirm;
    let asked: (string | undefined)[];
    beforeEach(() => {
        confirm = global.confirm;
        asked = [];
        global.confirm = (message) => {
            asked.push(message);
            return false;
        };
    });
    afterEach(() => {
        global.confirm = confirm;
    });

    it('inserts empty rows before a row, without a history record or an event', () => {
        const { worksheet, events } = load({ data: grid(3, 2) });

        expect(worksheet.shiftRows(1, 2, true)).to.equal(undefined);

        expect(worksheet.getData()).to.eql([
            ['A1', 'B1'],
            ['', ''],
            ['', ''],
            ['A2', 'B2'],
            ['A3', 'B3'],
        ]);
        expect(worksheet.rows.length).to.equal(5);
        expect(worksheet.history.records.length).to.equal(0);
        expect(dataEvents(events)).to.eql([]);
        expect(worksheet.parent.ignoreEvents).to.not.equal(true);
        expect(worksheet.history.ignore).to.equal(false);
        expect(worksheet.quiet).to.equal(false);
    });

    it('inserts rows after the last row when the position is the number of rows', () => {
        const { worksheet } = load({ data: grid(2, 2) });

        worksheet.shiftRows(2, 1, true);

        expect(worksheet.getData()).to.eql([
            ['A1', 'B1'],
            ['A2', 'B2'],
            ['', ''],
        ]);
    });

    it('deletes rows, without a history record or an event', () => {
        const { worksheet, events } = load({ data: grid(4, 2) });

        expect(worksheet.shiftRows(1, 2, false)).to.equal(undefined);

        expect(worksheet.getData()).to.eql([
            ['A1', 'B1'],
            ['A4', 'B4'],
        ]);
        expect(worksheet.history.records.length).to.equal(0);
        expect(dataEvents(events)).to.eql([]);
    });

    it('ignores allowInsertRow and allowDeleteRow, which apply to the user', () => {
        const { worksheet } = load({ data: grid(3, 1), allowInsertRow: false, allowDeleteRow: false });

        worksheet.shiftRows(0, 1, true);
        worksheet.shiftRows(3, 1, false);

        expect(worksheet.getData()).to.eql([[''], ['A1'], ['A2']]);
    });

    it('keeps the formula texts, because the source of the change also changed them', () => {
        const { worksheet } = load({
            data: [
                ['1', '=A2'],
                ['2', '=SUM(A1:A2)'],
            ],
        });

        worksheet.shiftRows(1, 1, true);
        expect(worksheet.getValueFromCoords(1, 0, false)).to.equal('=A2');
        expect(worksheet.getValueFromCoords(1, 2, false)).to.equal('=SUM(A1:A2)');

        worksheet.shiftRows(0, 1, false);
        expect(worksheet.getValueFromCoords(1, 1, false)).to.equal('=SUM(A1:A2)');
    });

    it('updates the formula texts of an insertRow call after a shift', () => {
        const { worksheet } = load({
            data: [
                ['1', '=A2'],
                ['2', ''],
            ],
        });

        worksheet.shiftRows(2, 1, true);
        worksheet.insertRow(1, 0, true);

        expect(worksheet.getValueFromCoords(1, 1, false)).to.equal('=A3');
    });

    it('removes a merged range that the change crosses without a question, and moves the others', () => {
        const { worksheet } = load({ data: grid(6, 2), mergeCells: { A1: [1, 3], B5: [1, 2] } });

        worksheet.shiftRows(1, 1, true);

        expect(asked).to.eql([]);
        expect(worksheet.getMerge()).to.eql({ B6: [1, 2] });

        worksheet.shiftRows(0, 1, false);
        expect(worksheet.getMerge()).to.eql({ B5: [1, 2] });
    });

    it('clears the search without a question', () => {
        const { worksheet } = load({ data: grid(3, 1), search: true });
        worksheet.search('A2');

        worksheet.shiftRows(0, 1, true);

        expect(asked).to.eql([]);
        expect(worksheet.getData()).to.eql([[''], ['A1'], ['A2'], ['A3']]);
    });

    it('keeps the selection on its cells', () => {
        const { worksheet } = load({ data: grid(4, 2) });
        worksheet.updateSelectionFromCoords(1, 2, 1, 3);

        worksheet.shiftRows(0, 2, true);
        expect(worksheet.selection.range).to.eql([1, 4, 1, 5]);

        worksheet.shiftRows(0, 1, false);
        expect(worksheet.selection.range).to.eql([1, 3, 1, 4]);
    });

    it('closes an editor in a deleted row without a save', () => {
        const { worksheet } = load({ data: grid(3, 1) });
        worksheet.openEditor(worksheet.records[1][0].element);
        (worksheet.editor.cell!.children[0] as HTMLInputElement).value = 'X';

        worksheet.shiftRows(1, 1, false);

        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.getData()).to.eql([['A1'], ['A3']]);
        expect(worksheet.history.records.length).to.equal(0);
    });

    it('keeps an editor in a row that moves', () => {
        const { worksheet } = load({ data: grid(3, 1) });
        const cell = worksheet.records[1][0].element;
        worksheet.openEditor(cell);

        worksheet.shiftRows(0, 1, true);

        expect(worksheet.editor.cell).to.equal(cell);
        expect(worksheet.editor.y).to.equal(2);
    });

    it('keeps the events, the history and the questions on after the change', () => {
        const { worksheet, events } = load({ data: grid(3, 1) });

        worksheet.shiftRows(0, 1, true);
        worksheet.insertRow(1, 0, true);

        expect(dataEvents(events)).to.include('oninsertrow');
        expect(worksheet.history.records.length).to.equal(1);
    });

    it('returns false for an argument that is not valid, and changes nothing', () => {
        const { worksheet } = load({ data: grid(2, 1) });

        expect(worksheet.shiftRows(-1, 1, true)).to.equal(false);
        expect(worksheet.shiftRows(3, 1, true)).to.equal(false);
        expect(worksheet.shiftRows(2, 1, false)).to.equal(false);
        expect(worksheet.shiftRows(0, 0, true)).to.equal(false);
        expect(worksheet.shiftRows(0.5, 1, true)).to.equal(false);
        expect(worksheet.getData()).to.eql([['A1'], ['A2']]);
    });

    it('returns false in the row source mode, where the application calls resetRows', () => {
        const { worksheet } = load({ minDimensions: [1, 1], tableOverflow: true, rowSource: { count: 2, row: (y) => [`${y}`] } });

        expect(worksheet.shiftRows(0, 1, true)).to.equal(false);
        expect(worksheet.shiftRows(0, 1, false)).to.equal(false);
    });

    it('restores the flags when the change throws', () => {
        const { worksheet } = load({ data: grid(2, 1) });
        const insert = worksheet.rowManager.insert;
        worksheet.rowManager.insert = () => {
            throw new Error('failed');
        };

        try {
            expect(() => worksheet.shiftRows(0, 1, true)).to.throw('failed');
        } finally {
            worksheet.rowManager.insert = insert;
        }
        expect(worksheet.parent.ignoreEvents).to.not.equal(true);
        expect(worksheet.history.ignore).to.equal(false);
        expect(worksheet.quiet).to.equal(false);
    });
});

describe('shiftColumns', () => {
    let confirm: typeof global.confirm;
    let asked: (string | undefined)[];
    beforeEach(() => {
        confirm = global.confirm;
        asked = [];
        global.confirm = (message) => {
            asked.push(message);
            return false;
        };
    });
    afterEach(() => {
        global.confirm = confirm;
    });

    it('inserts empty columns before a column, without a history record or an event', () => {
        const { worksheet, events } = load({ data: grid(2, 2) });

        expect(worksheet.shiftColumns(1, 2, true)).to.equal(undefined);

        expect(worksheet.getData()).to.eql([
            ['A1', '', '', 'B1'],
            ['A2', '', '', 'B2'],
        ]);
        expect(worksheet.headers.length).to.equal(4);
        expect(worksheet.history.records.length).to.equal(0);
        expect(dataEvents(events)).to.eql([]);
    });

    it('inserts columns after the last column when the position is the number of columns', () => {
        const { worksheet } = load({ data: grid(1, 2) });

        worksheet.shiftColumns(2, 1, true);

        expect(worksheet.getData()).to.eql([['A1', 'B1', '']]);
    });

    it('deletes columns, without a history record or an event', () => {
        const { worksheet, events } = load({ data: grid(1, 4) });

        expect(worksheet.shiftColumns(1, 2, false)).to.equal(undefined);

        expect(worksheet.getData()).to.eql([['A1', 'D1']]);
        expect(worksheet.history.records.length).to.equal(0);
        expect(dataEvents(events)).to.eql([]);
    });

    it('ignores allowInsertColumn and allowDeleteColumn, which apply to the user', () => {
        const { worksheet } = load({ data: grid(1, 2), allowInsertColumn: false, allowDeleteColumn: false });

        worksheet.shiftColumns(0, 1, true);
        worksheet.shiftColumns(2, 1, false);

        expect(worksheet.getData()).to.eql([['', 'A1']]);
    });

    it('keeps the formula texts, because the source of the change also changed them', () => {
        const { worksheet } = load({ data: [['1', '2', '=B1']] });

        worksheet.shiftColumns(1, 1, true);
        expect(worksheet.getValueFromCoords(3, 0, false)).to.equal('=B1');

        worksheet.shiftColumns(0, 1, false);
        expect(worksheet.getValueFromCoords(2, 0, false)).to.equal('=B1');
    });

    it('removes a merged range that the change crosses without a question, and moves the others', () => {
        const { worksheet } = load({ data: grid(2, 6), mergeCells: { A1: [3, 1], E2: [2, 1] } });

        worksheet.shiftColumns(1, 1, true);

        expect(asked).to.eql([]);
        expect(worksheet.getMerge()).to.eql({ F2: [2, 1] });
    });

    it('keeps the selection on its cells', () => {
        const { worksheet } = load({ data: grid(2, 4) });
        worksheet.updateSelectionFromCoords(2, 0, 3, 1);

        worksheet.shiftColumns(0, 1, true);

        expect(worksheet.selection.range).to.eql([3, 0, 4, 1]);
    });

    it('returns false for an argument that is not valid, and changes nothing', () => {
        const { worksheet } = load({ data: grid(1, 2) });

        expect(worksheet.shiftColumns(-1, 1, true)).to.equal(false);
        expect(worksheet.shiftColumns(3, 1, true)).to.equal(false);
        expect(worksheet.shiftColumns(2, 1, false)).to.equal(false);
        expect(worksheet.shiftColumns(0, 0, false)).to.equal(false);
        expect(worksheet.getData()).to.eql([['A1', 'B1']]);
    });
});
