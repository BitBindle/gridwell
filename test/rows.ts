import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('Use the rows method', () => {
    it('deleteRow and a row is removed', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody')!;
        let rows = table.children;
        let firstRow = rows[0];

        // Check that first row has the value of [1, 2, 3]
        expect(firstRow.children[1].innerHTML).to.include(1);
        expect(firstRow.children[2].innerHTML).to.include(2);
        expect(firstRow.children[3].innerHTML).to.include(3);

        instance[0].deleteRow(0);

        table = root.querySelector('tbody')!;
        rows = table.children;
        firstRow = rows[0];

        // Check that the value of the first row now is [4, 5, 6] since the first one got removed
        expect(firstRow.children[1].innerHTML).to.include(4);
        expect(firstRow.children[2].innerHTML).to.include(5);
        expect(firstRow.children[3].innerHTML).to.include(6);
    });

    it('insertRow and a row is added', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody')!;
        let rows = table.children;
        let firstRow = rows[0];

        // Check that first row has the value of [1, 2, 3]
        expect(firstRow.children[1].innerHTML).to.include(1);
        expect(firstRow.children[2].innerHTML).to.include(2);
        expect(firstRow.children[3].innerHTML).to.include(3);

        instance[0].insertRow([9, 9, 9], 0, 1);

        table = root.querySelector('tbody')!;
        rows = table.children;
        firstRow = rows[0];

        // Check that the value of the first row now is [9, 9, 9]
        expect(firstRow.children[1].innerHTML).to.include(9);
        expect(firstRow.children[2].innerHTML).to.include(9);
        expect(firstRow.children[3].innerHTML).to.include(9);
    });

    it('moveRow and the row is moved', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody')!;
        let rows = table.children;
        let firstRow = rows[0];
        let secondRow = rows[1];
        let A1 = firstRow.children[1];
        let A2 = secondRow.children[1];

        expect(A1.innerHTML).to.include(1);
        expect(A2.innerHTML).to.include(4);

        instance[0].moveRow(0, 1);

        table = root.querySelector('tbody')!;
        rows = table.children;
        firstRow = rows[0];
        secondRow = rows[1];
        A1 = firstRow.children[1];
        A2 = secondRow.children[1];

        expect(A1.innerHTML).to.include(4);
        expect(A2.innerHTML).to.include(1);
    });

    it('deleteRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].deleteRow(0);

        expect(rows[0].children[1].innerHTML).to.include(4);
        expect(rows[0].children[2].innerHTML).to.include(5);
        expect(rows[0].children[3].innerHTML).to.include(6);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(4);
        expect(rows[0].children[2].innerHTML).to.include(5);
        expect(rows[0].children[3].innerHTML).to.include(6);
    });

    it('insertRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].insertRow([9, 9, 9], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);
    });

    it('moveRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;

        expect(table.children[0].children[1].innerHTML).to.include(1);
        expect(table.children[1].children[1].innerHTML).to.include(4);

        instance[0].moveRow(0, 1);

        expect(table.children[0].children[1].innerHTML).to.include(4);
        expect(table.children[1].children[1].innerHTML).to.include(1);

        instance[0].undo();

        expect(table.children[0].children[1].innerHTML).to.include(1);
        expect(table.children[1].children[1].innerHTML).to.include(4);

        instance[0].redo();

        expect(table.children[0].children[1].innerHTML).to.include(4);
        expect(table.children[1].children[1].innerHTML).to.include(1);
    });
});

describe('Drag a row', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    // jsdom has no layout, so offsetX and offsetY are set: a press at the right edge of a row number starts a drag
    const mouse = (element: Element, type: string) => {
        const event = new window.MouseEvent(type, { bubbles: true, buttons: 1 });
        Object.defineProperty(event, 'offsetX', { value: 0 });
        Object.defineProperty(event, 'offsetY', { value: 0 });
        element.dispatchEvent(event);
    };

    it('ends the drag of the only row without an error', async () => {
        // Without row resize, the press at the edge is for a drag
        const worksheet = await load({ worksheets: [{ data: [['a']], rowResize: false }] });

        const errors: string[] = [];
        const onerror = (event: ErrorEvent) => {
            errors.push(event.message);
            // Keep jsdom from reporting the error to the console
            event.preventDefault();
        };
        window.addEventListener('error', onerror);

        const rowNumber = worksheet.rows[0].element.children[0];
        mouse(rowNumber, 'mousedown');
        expect(worksheet.pointer.dragging).to.not.equal(null);
        mouse(rowNumber, 'mouseup');

        window.removeEventListener('error', onerror);

        expect(errors).to.eql([]);
        expect(worksheet.pointer.dragging).to.equal(null);
        expect(worksheet.getData()).to.eql([['a']]);
    });
});

describe('Insert rows', () => {
    it('gives the events the new rows, and puts them before or after the row', () => {
        const events: [string, Parameters<NonNullable<Jspreadsheet.SpreadsheetOptions['oninsertrow']>>[1]][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
            onbeforeinsertrow: (instance, rows) => {
                events.push(['before', rows]);
            },
            oninsertrow: (instance, rows) => {
                events.push(['after', rows]);
            },
        })[0];

        worksheet.insertRow(['x', 'y'], 0, true);
        worksheet.insertRow(2, 2);

        expect(worksheet.getData()).to.eql([
            ['x', 'y'],
            ['a', 'b'],
            ['c', 'd'],
            ['', ''],
            ['', ''],
        ]);
        expect(events).to.eql([
            ['before', [{ row: 0, data: ['x', 'y'] }]],
            ['after', [{ row: 0, data: ['x', 'y'] }]],
            [
                'before',
                [
                    { row: 3, data: ['', ''] },
                    { row: 4, data: ['', ''] },
                ],
            ],
            [
                'after',
                [
                    { row: 3, data: ['', ''] },
                    { row: 4, data: ['', ''] },
                ],
            ],
        ]);
        expect(worksheet.rows.map((row) => row.y)).to.eql([0, 1, 2, 3, 4]);
        expect(Array.from(worksheet.tbody.children)).to.eql(worksheet.rows.map((row) => row.element));
        expect(worksheet.records.map((row) => row.map((cell) => cell.y))).to.eql([
            [0, 0],
            [1, 1],
            [2, 2],
            [3, 3],
            [4, 4],
        ]);
    });
});

describe('Insert rows in a worksheet without rows (#244)', () => {
    function create(options: Record<string, unknown> = {}) {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    allowDeletingAllRows: true,
                    data: [
                        ['1', 'a'],
                        ['2', 'b'],
                    ],
                    ...options,
                },
            ],
        })[0];
        worksheet.deleteRow(0, 2);
        return worksheet;
    }

    it('inserts a row, and the undo and the redo remove it and give it back', () => {
        const worksheet = create();

        worksheet.insertRow();
        expect(worksheet.getData()).to.eql([['', '']]);
        expect(Array.from(worksheet.tbody.children)).to.eql([worksheet.rows[0].element]);

        worksheet.undo();
        expect(worksheet.getData()).to.eql([]);
        expect(worksheet.tbody.children).to.have.length(0);

        worksheet.redo();
        expect(worksheet.getData()).to.eql([['', '']]);
        expect(Array.from(worksheet.tbody.children)).to.eql([worksheet.rows[0].element]);
    });

    it('inserts the rows at row 0 for each row number and direction', () => {
        const after = create();
        after.insertRow(2, 0);
        expect(after.rows.map((row) => row.y)).to.eql([0, 1]);
        expect(Array.from(after.tbody.children)).to.eql(after.rows.map((row) => row.element));

        const before = create();
        before.insertRow(['x', 'y'], 5, true);
        expect(before.getData()).to.eql([['x', 'y']]);
        expect(before.records[0].map((cell) => cell.y)).to.eql([0, 0]);
    });

    it('shows the new row with pagination', () => {
        const worksheet = create({ pagination: 3 });

        worksheet.insertRow();

        expect(Array.from(worksheet.tbody.children)).to.eql([worksheet.rows[0].element]);
    });
});

describe('Insert a row next to the rows that a filter shows', () => {
    it('shows a new row after a row that shows, and the redo shows the same rows', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ filters: true, data: [['a'], ['b'], ['a'], ['b']] }] })[0];
        worksheet.view.filters[0] = ['a'];
        worksheet.view.closeFilter(0);
        const attached = () => Array.from(worksheet.tbody.children).map((tr) => Number(tr.getAttribute('data-y')));
        expect(attached()).to.eql([0, 2]);

        worksheet.insertRow(1, 0);
        expect(attached()).to.eql([0, 1, 3]);

        worksheet.undo();
        expect(attached()).to.eql([0, 2]);

        worksheet.redo();
        expect(attached()).to.eql([0, 1, 3]);
    });
});

describe('Row heights, visibility and data', () => {
    it('getHeight returns the height of a row, and of the rows that have one', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [1, 2],
                        [3, 4],
                        [5, 6],
                    ],
                    rows: { 1: { height: '40px' } },
                },
            ],
        })[0];

        expect(worksheet.getHeight(1)).to.equal('40px');
        expect(worksheet.getHeight(0)).to.equal('');

        worksheet.setHeight(2, 50);

        expect(worksheet.getHeight(2)).to.equal('50px');
        expect(worksheet.getHeight()).to.eql([undefined, '40px', '50px']);
    });

    it('hideRow and showRow hide and show the rows', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [1, 2],
                        [3, 4],
                        [5, 6],
                    ],
                },
            ],
        })[0];

        worksheet.hideRow([0, 2]);

        expect(worksheet.rows.map((row) => row.element.style.display)).to.eql(['none', '', 'none']);

        worksheet.showRow(0);

        expect(worksheet.rows.map((row) => row.element.style.display)).to.eql(['', '', 'none']);

        worksheet.showRow([2]);

        expect(worksheet.rows.map((row) => row.element.style.display)).to.eql(['', '', '']);
    });

    it('getRowData and setRowData read and change the values of a row', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [1, 2, '=A1+B1'],
                        [4, 5, '=A2+B2'],
                    ],
                },
            ],
        })[0];

        expect(worksheet.getRowData(1)).to.eql([4, 5, '=A2+B2']);
        expect(worksheet.getRowData(1, true)).to.eql(['4', '5', '9']);

        // A null value leaves the cell as it is
        worksheet.setRowData(0, ['a', null, 7]);

        expect(worksheet.getRowData(0)).to.eql(['a', 2, 7]);
        expect(worksheet.getRowData(0, true)).to.eql(['a', '2', '7']);
        expect(worksheet.getValue('A1')).to.equal('a');
    });
});
