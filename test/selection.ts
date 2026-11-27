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

    it('uses the height of all the header rows for nested headers, because they stick below each other', () => {
        const worksheet = scrolledWorksheet([24, 26], [[{ title: 'Group', colspan: 2 }]]);

        up(worksheet, 20);

        expect(worksheet.content.scrollTop).to.equal(500 + 20 - 50);
    });
});

describe('Insert and delete keep the selection and the editor on their cells', () => {
    function create() {
        const data = Array.from({ length: 6 }, (row, y) => Array.from({ length: 6 }, (cell, x) => `${x}:${y}`));
        return jspreadsheet(root, { worksheets: [{ data }] })[0];
    }

    const range = (worksheet: Worksheet) => worksheet.selection.range?.map(Number) ?? null;
    const highlighted = (worksheet: Worksheet) => Array.from(worksheet.table.querySelectorAll('td.highlight')).map((td) => td.textContent);
    const input = (worksheet: Worksheet) => worksheet.editor.cell!.querySelector<HTMLInputElement | HTMLTextAreaElement>('input, textarea')!;

    it('moves the selection with its cells when rows or columns are inserted before it', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 3, 3);

        worksheet.insertRow(2, 1, true);
        expect(range(worksheet)).to.eql([2, 5, 3, 5]);
        expect(highlighted(worksheet)).to.eql(['2:3', '3:3']);
        expect(worksheet.selection.activeCell!.textContent).to.equal('2:3');

        worksheet.insertColumn(1, 0, true);
        expect(range(worksheet)).to.eql([3, 5, 4, 5]);
        expect(highlighted(worksheet)).to.eql(['2:3', '3:3']);
    });

    it('moves the selection with its cells when rows or columns before it are deleted', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 3, 4);

        worksheet.deleteRow(0, 2);
        expect(range(worksheet)).to.eql([2, 1, 3, 2]);
        worksheet.deleteColumn(0, 1);
        expect(range(worksheet)).to.eql([1, 1, 2, 2]);
        expect(highlighted(worksheet)).to.eql(['2:3', '3:3', '2:4', '3:4']);
    });

    it('grows the selection when the new rows split it', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 1, 2, 4);

        worksheet.insertRow(2, 2, true);

        expect(range(worksheet)).to.eql([2, 1, 2, 6]);
    });

    it('removes a selection that is inside the deleted rows', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 2, 2, 2);

        worksheet.deleteRow(1, 3);

        expect(range(worksheet)).to.equal(null);
        expect(highlighted(worksheet)).to.eql([]);
    });

    it('moves the selection back and forth with undo and redo', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 2, 3);

        worksheet.deleteRow(0, 1);
        worksheet.undo();
        expect(range(worksheet)).to.eql([2, 3, 2, 3]);
        worksheet.redo();
        expect(range(worksheet)).to.eql([2, 2, 2, 2]);

        worksheet.insertColumn(1, 0, true);
        expect(range(worksheet)).to.eql([3, 2, 3, 2]);
        worksheet.undo();
        expect(range(worksheet)).to.eql([2, 2, 2, 2]);
        expect(highlighted(worksheet)).to.eql(['2:3']);
    });

    it('keeps the editor open on its cell when a row or a column is inserted before it, and saves to that cell', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 2, 3);
        worksheet.openEditor(worksheet.records[3][2].element);

        worksheet.insertRow(1, 0, true);
        worksheet.insertColumn(1, 0, true);

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect([worksheet.editor.x, worksheet.editor.y]).to.eql([3, 4]);
        expect(worksheet.editor.cell).to.equal(worksheet.records[4][3].element);

        input(worksheet).value = 'typed';
        worksheet.closeEditor(worksheet.editor.cell!, true);
        expect(worksheet.getValueFromCoords(3, 4)).to.equal('typed');
    });

    it('closes the editor without a save when its row is deleted', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 2, 3);
        worksheet.openEditor(worksheet.records[3][2].element);
        input(worksheet).value = 'typed';

        worksheet.deleteRow(3, 1);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getData().flat()).to.not.include('typed');
        expect(worksheet.getValueFromCoords(2, 3)).to.equal('2:4');
    });

    it('closes the editor without a save when its column is deleted', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 3, 2, 3);
        worksheet.openEditor(worksheet.records[3][2].element);
        input(worksheet).value = 'typed';

        worksheet.deleteColumn(2, 1);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getData().flat()).to.not.include('typed');
    });

    // The spreadsheet adds its event handlers after a promise. jsdom gives the headers and the row numbers no size, so a
    // press is at their edge, which resizes or drags unless that is off.
    const load = (rows: number) =>
        new Promise<Worksheet>((resolve) => {
            const data = Array.from({ length: rows }, (row, y) => Array.from({ length: 3 }, (cell, x) => `${x}:${y}`));
            const options = { data, rowResize: false, rowDrag: false, columnResize: false, columnDrag: false };
            const instance = jspreadsheet(root, { worksheets: [options], onload: () => resolve(instance[0]) });
        });
    const press = (element: Element, options?: MouseEventInit) => {
        element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, ...options }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, buttons: 1 }));
    };
    const rowNumber = (worksheet: Worksheet, y: number) => worksheet.rows[y].element.children[0];

    it('moves a reversed selection, and a selection after rows or columns inserted after a cell', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(3, 4, 2, 2);

        worksheet.insertRow(1, 3, false);
        expect(range(worksheet)).to.eql([3, 5, 2, 2]);
        worksheet.insertColumn(1, 1, false);
        expect(range(worksheet)).to.eql([4, 5, 3, 2]);
        worksheet.deleteRow(0, 1);
        expect(range(worksheet)).to.eql([4, 4, 3, 1]);
        // The new row splits the selection, so its two empty cells are selected too
        expect(highlighted(worksheet)).to.eql(['2:2', '3:2', '2:3', '3:3', '', '', '2:4', '3:4']);
    });

    it('removes a selection with a cell in the deleted columns', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(1, 1, 3, 1);

        worksheet.deleteColumn(3, 2);

        expect(range(worksheet)).to.equal(null);
    });

    it('keeps the editor open when a press goes into it after an insert before it', async () => {
        const worksheet = await load(3);
        worksheet.updateSelectionFromCoords(0, 2, 0, 2);
        worksheet.openEditor(worksheet.records[2][0].element);

        worksheet.insertRow(1, 0, true);
        press(input(worksheet));

        expect(worksheet.editor.isOpen()).to.equal(true);
    });

    it('moves the first row number of a selection of whole rows, so that a Shift press extends it from that row', async () => {
        const worksheet = await load(6);

        press(rowNumber(worksheet, 2));
        worksheet.insertRow(1, 0, true);
        expect(worksheet.selection.row).to.equal(3);

        press(rowNumber(worksheet, 4), { shiftKey: true });
        expect(range(worksheet)).to.eql([0, 3, 2, 4]);

        worksheet.deleteRow(3, 1);
        expect(range(worksheet)).to.equal(null);
        expect(worksheet.selection.row).to.equal(null);
    });

    it('moves the first header of a selection of whole columns', async () => {
        const worksheet = await load(3);

        press(worksheet.headers[1]);
        worksheet.insertColumn(1, 0, true);
        expect(worksheet.selection.column).to.equal(2);

        worksheet.deleteColumn(0, 1);
        expect(worksheet.selection.column).to.equal(1);
        expect(range(worksheet)).to.eql([1, 0, 1, 2]);
    });

    it('forgets the first row number of a selection of whole rows when its row is deleted', async () => {
        const worksheet = await load(8);

        press(rowNumber(worksheet, 1));
        worksheet.updateSelectionFromCoords(0, 6, 0, 6);
        worksheet.deleteRow(0, 3);
        expect(worksheet.selection.row).to.equal(null);
        expect(range(worksheet)).to.eql([0, 3, 0, 3]);

        // Without the first row number, a Shift press extends from the first row of the selection
        press(rowNumber(worksheet, 2), { shiftKey: true });
        expect(range(worksheet)).to.eql([0, 3, 2, 2]);
    });
});

describe('The first row number or header of a selection of whole rows or columns', () => {
    // The spreadsheet adds its event handlers after a promise. jsdom gives the headers and the row numbers no size, so a
    // press is at their edge, which resizes or drags unless that is off.
    let questions: string[];
    const load = () =>
        new Promise<Worksheet>((resolve) => {
            questions = [];
            const data = Array.from({ length: 5 }, (row, y) => ['A', 'B', 'C'].map((column) => `${column}${y + 1}`));
            const options = { data, rowResize: false, rowDrag: false, columnResize: false, columnDrag: false };
            const confirm = (text: string) => {
                questions.push(text);
                return true;
            };
            const instance = jspreadsheet(root, { confirm, worksheets: [options], onload: () => resolve(instance[0]) });
        });
    const press = (element: Element, options?: MouseEventInit) => {
        element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, ...options }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, buttons: 1 }));
    };
    const key = (worksheet: Worksheet, which: number) =>
        worksheet.element.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, which, keyCode: which } as KeyboardEventInit));
    const rowNumber = (worksheet: Worksheet, y: number) => worksheet.rows[y].element.children[0];
    const range = (worksheet: Worksheet) => worksheet.selection.range?.map(Number) ?? null;
    const column = (worksheet: Worksheet, x: number) => worksheet.getData().map((row) => row[x]);

    it('clears the cell and keeps the row when Delete follows an arrow key after a press on a row number', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 1));
        worksheet.records[1][0].element.focus();

        key(worksheet, 40);
        key(worksheet, 46);

        expect(questions).to.eql([]);
        expect(worksheet.rows.length).to.equal(5);
        expect(column(worksheet, 0)).to.eql(['A1', 'A2', '', 'A4', 'A5']);
    });

    it('clears the cell and keeps the column when Delete follows updateSelectionFromCoords after a press on a header', async () => {
        const worksheet = await load();
        press(worksheet.headers[1]);
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);
        worksheet.records[0][2].element.focus();

        key(worksheet, 46);

        expect(questions).to.eql([]);
        expect(worksheet.getData()[0]).to.eql(['A1', 'B1', '']);
    });

    it('still deletes the rows of a selection of whole rows', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 1));
        press(rowNumber(worksheet, 2), { shiftKey: true });
        worksheet.records[1][0].element.focus();

        key(worksheet, 46);

        expect(questions).to.eql(['Are you sure to delete the selected rows?']);
        expect(column(worksheet, 0)).to.eql(['A1', 'A4', 'A5']);
    });

    it('extends from the current selection when a Shift press on a row number follows another selection (#226)', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 1));
        worksheet.updateSelectionFromCoords(0, 3, 0, 3);

        press(rowNumber(worksheet, 4), { shiftKey: true });

        expect(range(worksheet)).to.eql([0, 3, 2, 4]);
        expect(worksheet.selection.rowAnchor()).to.equal(3);
    });

    it('extends from the first column of a cell selection with a Shift press on a header', async () => {
        const worksheet = await load();
        worksheet.updateSelectionFromCoords(0, 2, 0, 2);

        press(worksheet.headers[2], { shiftKey: true });

        expect(range(worksheet)).to.eql([0, 0, 2, 4]);
        expect(worksheet.selection.columnAnchor()).to.equal(0);
    });

    it('clears all cells when Delete follows selectAll after a press on the first row number', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 0));
        worksheet.selectAll();
        worksheet.records[0][0].element.focus();

        key(worksheet, 46);

        expect(questions).to.eql([]);
        expect(worksheet.rows.length).to.equal(5);
        expect(
            worksheet
                .getData()
                .flat()
                .every((value) => value === '')
        ).to.equal(true);
    });
    it('extends from the current selection with a Ctrl press on a row number, as with a Shift press', async () => {
        const worksheet = await load();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);

        press(rowNumber(worksheet, 4), { ctrlKey: true });

        expect(range(worksheet)).to.eql([0, 1, 2, 4]);
    });

    it('keeps a selection of whole columns whole when a row moves, also on undo, so Delete still deletes the column', async () => {
        const worksheet = await load();
        press(worksheet.headers[1]);

        worksheet.moveRow(0, 4);
        expect(range(worksheet)).to.eql([1, 0, 1, 4]);
        worksheet.undo();
        expect(range(worksheet)).to.eql([1, 0, 1, 4]);
        expect(worksheet.selection.columnAnchor()).to.equal(1);

        worksheet.records[0][1].element.focus();
        key(worksheet, 46);
        expect(questions).to.eql(['Are you sure to delete the selected columns?']);
        expect(worksheet.getData()[0]).to.eql(['A1', 'C1']);
    });

    it('keeps a selection of whole rows whole when a column moves, and selectAll when a row moves', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 1));

        worksheet.moveColumn(2, 0);
        expect(range(worksheet)).to.eql([0, 1, 2, 1]);
        expect(worksheet.selection.rowAnchor()).to.equal(1);

        worksheet.selectAll();
        worksheet.moveRow(0, 2);
        expect(range(worksheet)).to.eql([0, 0, 2, 4]);
    });

    it('gives a selection of whole rows or columns the columns or rows that an insert adds after the last one', async () => {
        const worksheet = await load();
        press(rowNumber(worksheet, 1));

        worksheet.insertColumn();
        expect(range(worksheet)).to.eql([0, 1, 3, 1]);
        expect(worksheet.selection.rowAnchor()).to.equal(1);

        press(worksheet.headers[1]);
        worksheet.insertRow();
        expect(range(worksheet)).to.eql([1, 0, 1, 5]);
        expect(worksheet.selection.columnAnchor()).to.equal(1);

        // A cell selection that ends in the last column does not grow
        worksheet.updateSelectionFromCoords(0, 2, 3, 2);
        worksheet.insertColumn();
        expect(range(worksheet)).to.eql([0, 2, 3, 2]);
    });

    it('does not extend an old selection of whole columns with a drag after a press on the selection corner', async () => {
        const worksheet = await load();
        press(worksheet.headers[1]);
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);

        worksheet.corner.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1 }));
        worksheet.headers[2].dispatchEvent(new window.MouseEvent('mouseover', { bubbles: true, cancelable: true, buttons: 1 }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, buttons: 1 }));

        expect(range(worksheet)).to.eql([0, 0, 0, 0]);
    });
});

describe('moveRow and moveColumn keep the selection and the editor on their cells (#222)', () => {
    function create() {
        const data = Array.from({ length: 5 }, (row, y) => ['A', 'B', 'C', 'D'].map((column) => `${column}${y + 1}`));
        return jspreadsheet(root, { worksheets: [{ data }] })[0];
    }

    const range = (worksheet: Worksheet) => worksheet.selection.range?.map(Number) ?? null;
    const highlighted = (worksheet: Worksheet) => Array.from(worksheet.table.querySelectorAll('td.highlight')).map((td) => td.textContent);

    it('moves the selected cell with its row and its column', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);

        worksheet.moveRow(1, 3);
        expect(range(worksheet)).to.eql([1, 3, 1, 3]);
        worksheet.moveColumn(1, 0);
        expect(range(worksheet)).to.eql([0, 3, 0, 3]);
        expect(highlighted(worksheet)).to.eql(['B2']);
    });

    it('moves the selected cell when another row moves over it', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(1, 2, 1, 2);

        worksheet.moveRow(0, 4);
        expect(range(worksheet)).to.eql([1, 1, 1, 1]);
        worksheet.moveRow(4, 0);
        expect(range(worksheet)).to.eql([1, 2, 1, 2]);
        expect(highlighted(worksheet)).to.eql(['B3']);
    });

    it('keeps each end of a larger selection on its cell', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(0, 1, 1, 3);

        worksheet.moveRow(2, 4);

        expect(range(worksheet)).to.eql([0, 1, 1, 2]);
    });

    it('keeps the numbers of a selection when the row moves inside it', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(1, 1, 1, 3);

        worksheet.moveRow(1, 3);

        expect(range(worksheet)).to.eql([1, 1, 1, 3]);
    });

    it('moves the selection back and forth with undo and redo', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);

        worksheet.moveColumn(2, 3);
        worksheet.undo();
        expect(range(worksheet)).to.eql([2, 0, 2, 0]);
        worksheet.redo();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
    });

    it('keeps the editor on its cell, and saves to that cell', () => {
        const worksheet = create();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.openEditor(worksheet.records[1][1].element);

        worksheet.moveRow(1, 3);
        worksheet.moveColumn(0, 2);

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect([worksheet.editor.x, worksheet.editor.y]).to.eql([0, 3]);
        expect(worksheet.editor.cell).to.equal(worksheet.records[3][0].element);

        worksheet.editor.cell!.querySelector<HTMLTextAreaElement>('textarea')!.value = 'typed';
        worksheet.closeEditor(worksheet.editor.cell!, true);
        expect(worksheet.getValueFromCoords(0, 3)).to.equal('typed');
    });
});

describe('A press that does not select keeps the first row number or header (#250)', () => {
    // The spreadsheet adds its event handlers after a promise. jsdom gives the headers and the row numbers no size, so a
    // press is at their edge, which resizes or drags unless that is off.
    let questions: string[];
    const load = (options: Record<string, unknown> = {}) =>
        new Promise<Worksheet>((resolve) => {
            questions = [];
            const data = Array.from({ length: 3 }, (row, y) => ['A', 'B', 'C'].map((column) => `${column}${y + 1}`));
            const confirm = (text: string) => {
                questions.push(text);
                return true;
            };
            const instance = jspreadsheet(root, { confirm, worksheets: [{ data, ...options }], onload: () => resolve(instance[0]) });
        });
    const press = (element: Element, options?: MouseEventInit) => {
        element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, ...options }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, buttons: 1 }));
    };
    const deleteKey = (worksheet: Worksheet) =>
        worksheet.element.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, which: 46, keyCode: 46 } as KeyboardEventInit));
    const rowNumber = (worksheet: Worksheet, y: number) => worksheet.rows[y].element.children[0];
    const range = (worksheet: Worksheet) => worksheet.selection.range?.map(Number) ?? null;

    it('keeps the header of a selection of whole columns after a press that resizes a row, so Delete deletes the column', async () => {
        const worksheet = await load({ columnResize: false, columnDrag: false });
        press(worksheet.headers[1]);

        press(rowNumber(worksheet, 0));
        expect(range(worksheet)).to.eql([1, 0, 1, 2]);
        expect(worksheet.selection.columnAnchor()).to.equal(1);

        worksheet.records[0][1].element.focus();
        deleteKey(worksheet);
        expect(questions).to.eql(['Are you sure to delete the selected columns?']);
        expect(worksheet.getData()[0]).to.eql(['A1', 'C1']);
    });

    it('keeps the header after a refused drag of a merged row', async () => {
        const worksheet = await load({ columnResize: false, columnDrag: false, rowResize: false, mergeCells: { A2: [2, 1] } });
        press(worksheet.headers[2]);

        press(rowNumber(worksheet, 1));

        expect(range(worksheet)).to.eql([2, 0, 2, 2]);
        expect(worksheet.selection.columnAnchor()).to.equal(2);
    });

    it('keeps the row number of a selection of whole rows after a press that resizes a column', async () => {
        const worksheet = await load({ rowResize: false, rowDrag: false });
        press(rowNumber(worksheet, 1));

        press(worksheet.headers[0]);

        expect(range(worksheet)).to.eql([0, 1, 2, 1]);
        expect(worksheet.selection.rowAnchor()).to.equal(1);
    });

    it('drops the header when a press on a row number selects the row, and the row number when a press on a header selects', async () => {
        const worksheet = await load({ rowResize: false, rowDrag: false, columnResize: false, columnDrag: false });
        press(worksheet.headers[1]);

        press(rowNumber(worksheet, 2));
        expect(worksheet.selection.columnAnchor()).to.equal(null);
        expect(worksheet.selection.rowAnchor()).to.equal(2);

        press(worksheet.headers[0]);
        expect(worksheet.selection.rowAnchor()).to.equal(null);
        expect(worksheet.selection.columnAnchor()).to.equal(0);
    });
});

describe('Keys after the selection of a cell that does not show (#240)', () => {
    function create() {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [4, 8] }] })[0];
        worksheet.hideRow(5);
        worksheet.updateSelectionFromCoords(3, 7, 3, 7);
        worksheet.updateSelectionFromCoords(1, 5, 1, 5);
        return worksheet;
    }
    const range = (worksheet: Worksheet) => worksheet.selection.range?.map(Number) ?? null;

    it('moves down from the cell that has the tab stop', () => {
        const worksheet = create();
        expect(worksheet.selection.activeCell).to.equal(worksheet.records[6][1].element);
        expect(range(worksheet)).to.eql([1, 5, 1, 5]);

        worksheet.down();

        expect(range(worksheet)).to.eql([1, 7, 1, 7]);
        expect(worksheet.selection.activeCell).to.equal(worksheet.records[7][1].element);
    });

    it('moves up and right from the cell that has the tab stop', () => {
        const up = create();
        up.up();
        expect(range(up)).to.eql([1, 4, 1, 4]);

        const right = create();
        right.right();
        expect(range(right)).to.eql([2, 6, 2, 6]);
    });

    it('extends the selected range with Shift', () => {
        const worksheet = create();

        worksheet.down(true);

        expect(range(worksheet)).to.eql([1, 5, 1, 6]);
    });
});
