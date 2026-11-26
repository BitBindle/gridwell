import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('Selection', () => {
    it('isSelected tells whether a cell is in the selected range', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a', 'b', 'c']] }] })[0];

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);

        expect(worksheet.isSelected(1, 0)).to.equal(true);
        expect(worksheet.isSelected(2, 0)).to.equal(false);
    });

    it('isSelected returns false when no cell is selected', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a', 'b']] }] })[0];

        expect(worksheet.isSelected(0, 0)).to.equal(false);
    });
    it('selectAll does not change the selection that the undo history restores', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.setValue('A1', 'x');
        worksheet.selectAll();
        worksheet.undo();

        expect(worksheet.getValue('A1')).to.equal('a');
        expect(worksheet.getSelection()).to.eql([0, 0, 0, 0]);
    });

    it('an arrow key does not change the selection that the undo history restores', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.setValue('A1', 'x');
        worksheet.right();
        worksheet.undo();

        expect(worksheet.getValue('A1')).to.equal('a');
        expect(worksheet.getSelection()).to.eql([0, 0, 0, 0]);
    });

    it('selectAll keeps the selection when onbeforeselection cancels it', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.parent.config.onbeforeselection = () => false;
        worksheet.selectAll();

        expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
    });
});

describe('Selection with the mouse', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });
    const press = (element: Element, options?: MouseEventInit) =>
        element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, buttons: 1, ...options }));

    it('keeps the selected cell as numbers', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        press(worksheet.records[1][1].element);
        expect(worksheet.selection.range).to.eql([1, 1, 1, 1]);

        press(worksheet.records[2][2].element, { shiftKey: true });
        expect(worksheet.selection.range).to.eql([1, 1, 2, 2]);
    });

    it('extends the selection to column 0 and row 0 when the mouse moves over them', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        press(worksheet.records[1][1].element);
        worksheet.records[0][0].element.dispatchEvent(new window.MouseEvent('mouseover', { bubbles: true, buttons: 1 }));
        expect(worksheet.selection.range).to.eql([1, 1, 0, 0]);
    });

    it('sorts by column 0 after a double click on its header', async () => {
        const worksheet = await load({ worksheets: [{ data: [['b'], ['a']] }] });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
        expect(worksheet.getData()).to.eql([['a'], ['b']]);
    });

    it('gives numbers to onbeforepaste after a mouse selection', async () => {
        const coords: (number | string)[] = [];
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            onbeforepaste: (instance, data, x, y) => {
                coords.push(x, y);
            },
        });

        press(worksheet.records[1][2].element);
        worksheet.paste(worksheet.selection.range![0], worksheet.selection.range![1], 'a');

        expect(coords).to.eql([2, 1]);
        expect(worksheet.getValue('C2')).to.equal('a');
    });
});

describe('The selection of a worksheet', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b', 'c'],
                        ['d', 'e', 'f'],
                    ],
                },
            ],
        })[0];

    it('has no range before a cell is selected', () => {
        const worksheet = create();

        expect(worksheet.selection.range).to.equal(null);
        expect(worksheet.selection.bounds).to.equal(null);
        expect(worksheet.selection.highlighted).to.eql([]);
    });

    it('keeps the range as given and the bounds sorted', () => {
        const worksheet = create();

        worksheet.updateSelectionFromCoords(2, 1, 1, 0);

        expect(worksheet.selection.range).to.eql([2, 1, 1, 0]);
        expect(worksheet.selection.bounds).to.eql([1, 0, 2, 1]);
    });

    it('keeps the records of the highlighted cells', () => {
        const worksheet = create();

        worksheet.updateSelectionFromCoords(1, 0, 2, 1);

        const cells = worksheet.selection.highlighted.map((record) => [record.x, record.y]);
        expect(cells).to.have.deep.members([
            [1, 0],
            [2, 0],
            [1, 1],
            [2, 1],
        ]);
        expect(worksheet.selection.highlighted[0].element).to.equal(worksheet.getCell('B1'));
    });

    it('has no range and no highlighted cells after resetSelection', () => {
        const worksheet = create();

        worksheet.updateSelectionFromCoords(0, 0, 1, 1);
        worksheet.resetSelection();

        expect(worksheet.selection.range).to.equal(null);
        expect(worksheet.selection.highlighted).to.eql([]);
    });

    it('moves the active cell with the selection', () => {
        const worksheet = create();

        worksheet.updateSelectionFromCoords(2, 1, 2, 1);
        expect(worksheet.selection.activeCell).to.equal(worksheet.getCell('C2'));

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);
        expect(worksheet.selection.activeCell).to.equal(worksheet.getCell('A1'));
    });

    it('clearFill removes the fill classes and empties the fill cells', () => {
        const worksheet = create();
        const cell = worksheet.getCell('A2');
        cell.classList.add('selection', 'selection-left', 'selection-bottom');
        worksheet.selection.fill.push(cell);

        worksheet.selection.clearFill();

        expect(worksheet.selection.fill).to.eql([]);
        expect(cell.className).to.not.match(/selection/);
    });

    it('is not shared between worksheets', async () => {
        const worksheets = jspreadsheet(root, { worksheets: [{ data: [['a']] }, { data: [['b']] }] });
        await new Promise((resolve) => setTimeout(resolve));

        expect(worksheets[0].selection).to.not.equal(worksheets[1].selection);

        worksheets[0].updateSelectionFromCoords(0, 0, 0, 0);

        expect(worksheets[1].selection.range).to.equal(null);
    });
});

describe('The pointer of a worksheet', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });
    const mouse = (element: EventTarget, type: string, options?: MouseEventInit) =>
        element.dispatchEvent(new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...options }));

    it('has no gesture before a press', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        expect(worksheet.pointer.resizing).to.equal(null);
        expect(worksheet.pointer.dragging).to.equal(null);
        expect(worksheet.pointer.cursor).to.equal(null);
        expect(worksheet.pointer.corner).to.equal(false);
        expect(worksheet.selection.column).to.equal(null);
        expect(worksheet.selection.row).to.equal(null);
        // The fields that the pointer and the selection replace
        expect(worksheet).to.not.have.property('resizing');
        expect(worksheet).to.not.have.property('dragging');
        expect(worksheet).to.not.have.property('selectedCorner');
        expect(worksheet).to.not.have.property('selectedHeader');
        expect(worksheet).to.not.have.property('selectedRow');
    });

    it('keeps the drag of the selection corner until the button goes up', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(document, 'mouseup');
        mouse(worksheet.corner, 'mousedown');

        expect(worksheet.pointer.corner).to.equal(true);

        mouse(document, 'mouseup');

        expect(worksheet.pointer.corner).to.equal(false);
    });

    it('fills the empty cells below the selection on a double click on the corner', async () => {
        const data = [
            ['1', 'a'],
            ['2', 'b'],
            ['', ''],
            ['', 'c'],
        ];
        const worksheet = await load({ worksheets: [{ data }] });

        worksheet.updateSelectionFromCoords(0, 0, 1, 1);
        mouse(worksheet.corner, 'dblclick');

        // The fill repeats the selected rows, and stops at a cell that has a value
        expect(worksheet.getData()).to.deep.equal([
            ['1', 'a'],
            ['2', 'b'],
            ['1', 'a'],
            ['2', 'c'],
        ]);
        expect(worksheet.history.records.length).to.equal(1);
    });

    it('stops the fill of a double click on the corner at a cell that holds 0 or false (issue #117)', async () => {
        const worksheet = await load({
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['', ''],
                        [0, false],
                        ['', ''],
                    ],
                },
            ],
        });

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);
        mouse(worksheet.corner, 'dblclick');

        expect(worksheet.getData()).to.deep.equal([
            ['a', 'b'],
            ['a', 'b'],
            [0, false],
            ['', ''],
        ]);
    });

    it('keeps the header and the row number that a press selected first', async () => {
        // jsdom gives the headers no size, so a press is at their edge, which resizes or drags unless that is off
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3], columnResize: false, columnDrag: false, rowResize: false, rowDrag: false }] });

        mouse(worksheet.headers[1], 'mousedown');
        mouse(document, 'mouseup');

        expect(worksheet.selection.column).to.equal(1);
        expect(worksheet.selection.range).to.eql([1, 0, 1, 2]);

        mouse(worksheet.rows[2].element.children[0], 'mousedown');
        mouse(document, 'mouseup');

        expect(worksheet.selection.row).to.equal(2);
        expect(worksheet.selection.column).to.equal(null);
        expect(worksheet.selection.range).to.eql([0, 2, 2, 2]);
    });
});

describe('A selection of hidden rows', () => {
    it('gives onbeforeselection and onselection the selected rows as the top and bottom (issue #59)', () => {
        const calls: (string | number)[][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 3] }],
            // A block, so that the handler returns nothing: onbeforeselection returns false or nothing
            onbeforeselection: (instance, left, top, right, bottom) => {
                calls.push(['before', left, top, right, bottom]);
            },
            onselection: (instance, left, top, right, bottom) => calls.push(['selection', left, top, right, bottom]),
        })[0];

        worksheet.hideRow(1);
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);

        expect(calls).to.eql([
            ['before', 1, 1, 1, 1],
            ['selection', 1, 1, 1, 1],
        ]);
        expect(worksheet.selection.highlighted).to.eql([]);
    });
});

describe('A move that onbeforeselection cancels', () => {
    const moves: ['up' | 'down' | 'left' | 'right' | 'first' | 'last', [number, number, number, number]][] = [
        ['up', [1, 1, 1, 1]],
        ['down', [1, 1, 1, 1]],
        ['left', [1, 1, 1, 1]],
        ['right', [1, 1, 1, 1]],
        ['first', [1, 1, 1, 1]],
        ['last', [1, 1, 1, 1]],
    ];

    for (const [move, start] of moves) {
        for (const shiftKey of [false, true]) {
            it(`keeps the selected range with ${move}${shiftKey ? ' and Shift' : ''} (issue #71)`, () => {
                const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 3] }] })[0];

                worksheet.updateSelectionFromCoords(...start);
                worksheet.parent.config.onbeforeselection = () => false;

                worksheet[move](shiftKey, false);

                expect(worksheet.getSelection()).to.eql(start);
                expect(worksheet.selection.highlighted.map((cell) => [cell.x, cell.y])).to.eql([[1, 1]]);
            });
        }
    }

    it('moves from the selected range with the next key', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 3] }] })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.parent.config.onbeforeselection = () => false;
        worksheet.right();
        // @ts-expect-error -- the test removes the handler with null
        worksheet.parent.config.onbeforeselection = null;
        worksheet.right();

        expect(worksheet.getSelection()).to.eql([1, 0, 1, 0]);
    });
});

describe('The scroll to the active cell below the header', () => {
    function rect(left: number, top: number, width: number, height: number) {
        // The rectangle has no toJSON, which the code does not call
        return () => ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top }) as DOMRect;
    }

    // The content is 300 pixels high and scrolls 500 pixels. Each header row has the given height.
    function scrolledWorksheet(headerRowHeights: number[], nestedHeaders?: { title: string; colspan: number }[][]) {
        const data: string[][] = [];
        for (let j = 0; j < 60; j++) {
            data.push([`A${j + 1}`, `B${j + 1}`]);
        }
        const worksheet = jspreadsheet(root, { worksheets: [{ data, nestedHeaders, tableOverflow: true }] })[0];

        Object.defineProperty(worksheet.content, 'scrollTop', { value: 500, writable: true });
        worksheet.content.getBoundingClientRect = rect(0, 0, 400, 300);
        // The first row starts below the header rows and the top border of the table
        const header = headerRowHeights.reduce((sum, height) => sum + height, 0);
        worksheet.tbody.getBoundingClientRect = rect(0, 1 + header - 500, 400, 1200);
        Array.from(worksheet.thead.children).forEach((row, i) => {
            Object.defineProperty(row, 'offsetHeight', { value: headerRowHeights[i] });
        });

        return worksheet;
    }

    function up(worksheet: Worksheet, cellTop: number) {
        worksheet.records[30][0].element.getBoundingClientRect = rect(50, cellTop, 100, 20);
        worksheet.updateSelectionFromCoords(0, 31, 0, 31);
        worksheet.up();
        expect(worksheet.selection.range!.map(Number)).to.eql([0, 30, 0, 30]);
    }

    it('ArrowUp puts the top edge of a cell that the header row covers in part below the header row', () => {
        const worksheet = scrolledWorksheet([40]);

        up(worksheet, 35);

        expect(worksheet.content.scrollTop).to.equal(500 + 35 - 40);
    });

    it('ArrowUp puts the top edge of a cell above the view below the header row', () => {
        const worksheet = scrolledWorksheet([40]);

        up(worksheet, -20);

        expect(worksheet.content.scrollTop).to.equal(500 - 20 - 40);
    });

    it('keeps the scroll position when ArrowUp selects a cell below the header row', () => {
        const worksheet = scrolledWorksheet([40]);

        up(worksheet, 40);

        expect(worksheet.content.scrollTop).to.equal(500);
    });

    it('ArrowRight puts the top edge of a cell that the header row covers in part below the header row', () => {
        const worksheet = scrolledWorksheet([40]);

        worksheet.records[30][1].element.getBoundingClientRect = rect(150, 35, 100, 20);
        worksheet.updateSelectionFromCoords(0, 30, 0, 30);
        worksheet.right();

        expect(worksheet.selection.range!.map(Number)).to.eql([1, 30, 1, 30]);
        expect(worksheet.content.scrollTop).to.equal(500 + 35 - 40);
    });

    it('scrolls to the top when ArrowUp selects a cell of the first row, so all the header rows show', () => {
        const worksheet = scrolledWorksheet([24, 26], [[{ title: 'Group', colspan: 2 }]]);

        // The first row starts 51 pixels below the top of the content when the content does not scroll
        worksheet.content.scrollTop = 40;
        worksheet.tbody.getBoundingClientRect = rect(0, 51 - 40, 400, 1200);
        worksheet.records[0][0].element.getBoundingClientRect = rect(50, 51 - 40, 100, 20);
        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        worksheet.up();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 0, 0, 0]);
        expect(worksheet.content.scrollTop).to.equal(0);
    });

    it('uses the highest header row for nested headers, because each header row sticks at the top', () => {
        const worksheet = scrolledWorksheet([24, 26], [[{ title: 'Group', colspan: 2 }]]);

        up(worksheet, 20);

        expect(worksheet.content.scrollTop).to.equal(500 + 20 - 26);
    });
});
