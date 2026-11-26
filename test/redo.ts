import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

describe('Use the redo method', () => {
    it('.undo', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');
        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('CE');

        instance[0].setValueFromCoords(0, 0, 'New Value');
        instance[0].setValueFromCoords(1, 0, 'TESTING');

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('TESTING');

        instance[0].undo();

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].undo();

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');
    });

    it('.redo after undo something', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');
        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('CE');

        instance[0].setValueFromCoords(0, 0, 'New Value');
        instance[0].setValueFromCoords(1, 0, 'TESTING');

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('TESTING');

        instance[0].undo();

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].undo();

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].redo();

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].redo();

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('TESTING');
    });
});

describe('The undo history', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                },
            ],
        })[0];

    it('keeps a record of each change', () => {
        const worksheet = create();

        expect(worksheet.history.records).to.deep.equal([]);
        expect(worksheet.history.index).to.equal(-1);
        expect(worksheet.history.canUndo()).to.equal(false);
        expect(worksheet.history.canRedo()).to.equal(false);

        worksheet.setValueFromCoords(0, 0, 'New Value');
        worksheet.setValueFromCoords(1, 0, 'TESTING');

        expect(worksheet.history.records.length).to.equal(2);
        expect(worksheet.history.records[0].action).to.equal('setValue');
        expect(worksheet.history.index).to.equal(1);
        expect(worksheet.history.canUndo()).to.equal(true);
        expect(worksheet.history.canRedo()).to.equal(false);
    });

    it('moves through the records with undo and redo', () => {
        const worksheet = create();

        worksheet.setValueFromCoords(0, 0, 'New Value');
        worksheet.setValueFromCoords(1, 0, 'TESTING');

        worksheet.undo();
        worksheet.undo();

        expect(worksheet.history.index).to.equal(-1);
        expect(worksheet.history.canUndo()).to.equal(false);
        expect(worksheet.history.canRedo()).to.equal(true);
        // Undo records nothing itself
        expect(worksheet.history.records.length).to.equal(2);

        worksheet.redo();

        expect(worksheet.history.index).to.equal(0);
        expect(worksheet.history.canUndo()).to.equal(true);
        expect(worksheet.history.canRedo()).to.equal(true);
    });

    it('drops the undone records when a new change is made', () => {
        const worksheet = create();

        worksheet.setValueFromCoords(0, 0, 'New Value');
        worksheet.setValueFromCoords(1, 0, 'TESTING');
        worksheet.undo();
        worksheet.setValueFromCoords(0, 1, 'Other');

        expect(worksheet.history.records.length).to.equal(2);
        expect((worksheet.history.records[1] as Extract<(typeof worksheet.history.records)[number], { action: 'setValue' }>).records[0].value).to.equal(
            'Other'
        );
        expect(worksheet.history.canRedo()).to.equal(false);
    });

    it('gives undo and redo events the record', () => {
        const events: [string, Jspreadsheet.HistoryRecord | undefined][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['Hello']] }],
            onundo: (instance, record) => events.push(['undo', record]),
            onredo: (instance, record) => events.push(['redo', record]),
        })[0];

        worksheet.setValueFromCoords(0, 0, 'New Value');
        const record = worksheet.history.records[0];

        worksheet.undo();
        worksheet.redo();
        worksheet.redo();

        expect(events).to.deep.equal([
            ['undo', record],
            ['redo', record],
            ['redo', undefined],
        ]);
    });

    it('keeps no records while ignore is set', () => {
        const worksheet = create();

        worksheet.history.ignore = true;
        worksheet.setValueFromCoords(0, 0, 'New Value');
        worksheet.history.ignore = false;

        expect(worksheet.history.records.length).to.equal(0);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('New Value');

        // Undo and redo restore the flag
        worksheet.setValueFromCoords(0, 0, 'Again');
        worksheet.undo();
        worksheet.redo();
        expect(worksheet.history.ignore).to.equal(false);
    });

    it('is cleared when the data is replaced', () => {
        const worksheet = create();

        worksheet.setValueFromCoords(0, 0, 'New Value');
        worksheet.setData([['A', 'B']]);

        expect(worksheet.history.records).to.deep.equal([]);
        expect(worksheet.history.index).to.equal(-1);
    });

    it('is not shared between worksheets', async () => {
        const worksheets = jspreadsheet(root, {
            worksheets: [{ data: [['A']] }, { data: [['B']] }],
        });

        // The worksheets after the first are built asynchronously
        await new Promise((resolve) => setTimeout(resolve));

        worksheets[0].setValueFromCoords(0, 0, 'New Value');

        expect(worksheets[0].history.records.length).to.equal(1);
        expect(worksheets[1].history.records.length).to.equal(0);
        expect(worksheets[0].history).to.not.equal(worksheets[1].history);
    });
});

describe('Undo and redo the changes to columns, rows and the order', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                        [7, 8, 9],
                    ],
                    columns: [{ width: 100 }, { width: 100 }, { width: 100 }],
                },
            ],
        })[0];

    const original = [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
    ];

    it('deleteColumn', () => {
        const worksheet = create();

        worksheet.deleteColumn(1);
        const deleted = [
            [1, 3],
            [4, 6],
            [7, 9],
        ];
        expect(worksheet.getData()).to.eql(deleted);

        worksheet.undo();
        expect(worksheet.getData()).to.eql(original);
        expect(worksheet.headers.length).to.equal(3);

        worksheet.redo();
        expect(worksheet.getData()).to.eql(deleted);
        expect(worksheet.headers.length).to.equal(2);
    });

    it('moveColumn', () => {
        const worksheet = create();

        worksheet.moveColumn(0, 2);
        const moved = [
            [2, 3, 1],
            [5, 6, 4],
            [8, 9, 7],
        ];
        expect(worksheet.getData()).to.eql(moved);

        worksheet.undo();
        expect(worksheet.getData()).to.eql(original);

        worksheet.redo();
        expect(worksheet.getData()).to.eql(moved);
    });

    it('orderBy', () => {
        const worksheet = create();

        worksheet.orderBy(0, 1);
        const ordered = [
            [7, 8, 9],
            [4, 5, 6],
            [1, 2, 3],
        ];
        expect(worksheet.getData()).to.eql(ordered);

        worksheet.undo();
        expect(worksheet.getData()).to.eql(original);

        worksheet.redo();
        expect(worksheet.getData()).to.eql(ordered);
    });

    it('setWidth', () => {
        const worksheet = create();

        worksheet.setWidth(0, 200);
        expect(worksheet.getWidth(0)).to.equal(200);

        worksheet.undo();
        expect(worksheet.getWidth(0)).to.equal(100);

        worksheet.redo();
        expect(worksheet.getWidth(0)).to.equal(200);
    });

    it('setHeight', () => {
        const worksheet = create();

        // The old height is the height that the browser measures. jsdom has no layout.
        // The test needs only the height
        worksheet.rows[1].element.getBoundingClientRect = () => ({ height: 25 } as DOMRect);

        worksheet.setHeight(1, 50);
        expect(worksheet.getHeight(1)).to.equal('50px');

        worksheet.undo();
        expect(worksheet.getHeight(1)).to.equal('25px');

        worksheet.redo();
        expect(worksheet.getHeight(1)).to.equal('50px');
    });
});
