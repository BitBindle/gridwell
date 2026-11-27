import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('Use the columns method', () => {
    it('insertColumn and column is inserted in the position 0', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(9);

        // Insert [3, 3] column in the first column
        instance[0].insertColumn([3, 3], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(3);

        // Insert [2, 2] column in the first column
        instance[0].insertColumn([2, 2], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(2);
    });

    it('insertColumn and column is inserted in the position 1', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[2].innerHTML).to.include(9);

        // Insert [3, 3] column in the second column
        instance[0].insertColumn([3, 3], 0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(3);

        // Insert [2, 2] column in the second column
        instance[0].insertColumn([2, 2], 0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(2);
    });

    it('deleteColumn and column is removed in the given index', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 9, 9, 9],
                        [3, 6, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(3);
        expect(rows[0].children[2].innerHTML).to.include(6);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(6);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);
    });

    it('deleteColumn and multiple column are removed starting from the given index', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 12, 15, 18],
                        [3, 6, 9, 12, 15, 18],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(3);
        expect(rows[0].children[2].innerHTML).to.include(6);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first two columns
        instance[0].deleteColumn(0, 2);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(12);
        expect(rows[0].children[3].innerHTML).to.include(15);

        // Delete first two columns
        instance[0].deleteColumn(0, 2);

        expect(rows[0].children[1].innerHTML).to.include(15);
        expect(rows[0].children[2].innerHTML).to.include(18);
    });

    it('hideColumn and showColumn', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 12, 15, 18],
                        [3, 6, 9, 12, 15, 18],
                    ],
                },
            ],
        });

        const table = root.querySelector('thead')!;
        const headers = table.children[0].children;

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).not.to.include('none');

        // Hides the first column 'A'
        instance[0].hideColumn(0);

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).to.include('none');
        expect(headers[2].innerHTML).to.include('B');
        expect(window.getComputedStyle(headers[2]).display).not.to.include('none');

        // Shows the column that was hidden
        instance[0].showColumn(0);

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).not.to.include('none');
        expect(headers[2].innerHTML).to.include('B');
        expect(window.getComputedStyle(headers[2]).display).not.to.include('none');
    });

    it('insertColumn history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(9);

        // Insert [3, 3] column in the first column
        instance[0].insertColumn([3, 3], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(3);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(9);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(3);
    });

    it('deleteColumn history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3, 4],
                        [5, 6, 7, 8],
                    ],
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(2);
        expect(rows[0].children[2].innerHTML).to.include(3);
        expect(rows[0].children[3].innerHTML).to.include(4);

        // @ts-expect-error -- the test passes an argument that undo does not take
        instance[0].undo(0);

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        // @ts-expect-error -- the test passes an argument that redo does not take
        instance[0].redo(0);

        expect(rows[0].children[1].innerHTML).to.include(2);
        expect(rows[0].children[2].innerHTML).to.include(3);
        expect(rows[0].children[3].innerHTML).to.include(4);
    });

    it('hideColumn keeps a merged range visible when it covers a column that is not hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Item', '', 'x', ''],
                        ['', '', '', ''],
                        ['', '', 'y', ''],
                    ],
                    mergeCells: { A1: [2, 1], C3: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.hideColumn(2);

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[2][2].element.style.display).to.equal('');
        expect(worksheet.records[1][1].element.style.display).to.equal('none');
        expect(worksheet.records[1][2].element.style.display).to.equal('none');

        // The visible first cell spans only the columns that are not hidden, so the next cells keep their columns
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('1');
        expect(worksheet.records[2][2].element.getAttribute('colspan')).to.equal('1');

        worksheet.showColumn([1, 2]);

        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('2');
        expect(worksheet.records[2][2].element.getAttribute('colspan')).to.equal('2');
    });

    it('hideColumn hides a merged range when each column that it covers is hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Item', '', 'x'],
                        ['', '', ''],
                    ],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        worksheet.hideColumn([0, 1]);

        expect(worksheet.records[0][0].element.style.display).to.equal('none');
    });

    it('setMerge gives the first cell only the columns that are not hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2] }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.setMerge('A1', 3, 1);

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('2');

        worksheet.showColumn(1);

        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('3');
    });

    it('keeps a merged range visible when one of its columns has the type hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['Item', 'b', 'c']],
                    columns: [{ type: 'hidden' }, {}, {}],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('1');
    });

    it('removeMerge keeps the cells of a hidden column hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: [['Item', '', 'x', 'y']], mergeCells: { A1: [2, 1] } }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(0);
        worksheet.removeMerge('A1');

        expect(worksheet.records[0][0].element.style.display).to.equal('none');
        expect(worksheet.records[0][1].element.style.display).to.equal('');

        worksheet.setMerge('A1', 2, 1);
        worksheet.hideColumn(1);
        worksheet.showColumn(0);
        worksheet.removeMerge('A1');

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][1].element.style.display).to.equal('none');
    });

    it('selects each column of a merged range when one of its columns is hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], mergeCells: { A1: [3, 1] } }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);

        expect(worksheet.headers.map((header) => header.classList.contains('selected'))).to.eql([true, false, true, false]);

        worksheet.resetSelection();

        expect(worksheet.headers.map((header) => header.classList.contains('selected'))).to.eql([false, false, false, false]);
    });

    describe('A merged range whose first column has the type hidden', () => {
        const source = [
            { id: '1', name: 'One' },
            { id: '2', name: 'Two' },
        ];

        it('shows and edits the first cell with the options of the first visible column', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '', 'x']],
                        columns: [{ type: 'hidden' }, { type: 'dropdown', source }, { type: 'text' }],
                        mergeCells: { A1: [2, 1] },
                    },
                ],
            })[0];

            const cell = worksheet.records[0][0].element;
            expect(cell.style.display).to.equal('');
            expect(cell.classList.contains('jss_dropdown')).to.equal(true);
            expect(cell.textContent).to.equal('One');

            worksheet.openEditor(cell);
            expect(cell.querySelector('.jdropdown')).to.not.equal(null);
            worksheet.closeEditor(cell, false);

            worksheet.setValueFromCoords(0, 0, '2');
            expect(cell.textContent).to.equal('Two');
            expect(worksheet.options.data[0][0]).to.equal('2');
        });

        it('saves the text of the editor to the first cell', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a', '']],
                        columns: [{ type: 'hidden' }, { type: 'text' }],
                        mergeCells: { A1: [2, 1] },
                    },
                ],
            })[0];

            const cell = worksheet.records[0][0].element;
            worksheet.openEditor(cell);
            (cell.children[0] as HTMLInputElement).value = 'b';
            worksheet.closeEditor(cell, true);

            expect(worksheet.options.data[0][0]).to.equal('b');
            expect(cell.textContent).to.equal('b');
        });

        it('renders the first cell again when setMerge, hideColumn, showColumn and removeMerge change the first visible column', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '', 'x']],
                        columns: [{ type: 'hidden' }, { type: 'dropdown', source }, { type: 'text' }],
                    },
                ],
            })[0];

            const cell = worksheet.records[0][0].element;

            worksheet.setMerge('A1', 3, 1);
            expect(cell.textContent).to.equal('One');

            worksheet.hideColumn(1);
            expect(cell.textContent).to.equal('1');
            expect(cell.classList.contains('jss_dropdown')).to.equal(false);

            worksheet.showColumn(1);
            expect(cell.textContent).to.equal('One');

            worksheet.removeMerge('A1');
            expect(cell.style.display).to.equal('none');
            expect(cell.textContent).to.equal('1');
            expect(cell.classList.contains('jss_dropdown')).to.equal(false);
        });

        it('keeps the readOnly state and the style of the first cell when hideColumn and showColumn change the first visible column', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '', 'x', 'y']],
                        columns: [{ type: 'hidden' }, { type: 'text' }, { type: 'text' }, { type: 'text' }],
                        mergeCells: { A1: [3, 1] },
                        style: { A1: 'text-align: left; white-space: nowrap;' },
                    },
                ],
            })[0];

            worksheet.setReadOnly('A1', true);

            worksheet.hideColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(true);
            expect(worksheet.records[0][0].element.style.textAlign.trim()).to.equal('left');
            expect(worksheet.records[0][0].element.style.whiteSpace.trim()).to.equal('nowrap');

            worksheet.showColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(true);
            expect(worksheet.records[0][0].element.style.textAlign.trim()).to.equal('left');
        });

        it('keeps the value of the hidden column when the first visible column is a checkbox column', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['id-7', '', 'x']],
                        columns: [{ type: 'hidden' }, { type: 'checkbox' }, { type: 'text' }],
                        mergeCells: { A1: [2, 1] },
                    },
                ],
            })[0];

            expect(worksheet.getValueFromCoords(0, 0)).to.equal('id-7');
            expect(worksheet.records[0][0].element.querySelector('input')).to.equal(null);
        });

        it('adds a new line with Alt+Enter when the first visible column has wordWrap', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a', '', 'x']],
                        columns: [{ type: 'hidden' }, { type: 'text', wordWrap: true }, { type: 'text' }],
                        mergeCells: { A1: [2, 1] },
                    },
                ],
            })[0];

            const cell = worksheet.records[0][0].element;
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            worksheet.openEditor(cell);
            const editor = cell.children[0] as HTMLTextAreaElement;
            editor.selectionStart = editor.selectionEnd = editor.value.length;
            editor.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, key: 'Enter', keyCode: 13, which: 13, altKey: true }));

            expect(worksheet.editor.cell).to.not.equal(null);
            expect(editor.value).to.equal('a\n');
            worksheet.closeEditor(cell, false);
        });

        it('uses the readOnly and align options of the first visible column when setMerge, hideColumn and showColumn change that column', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['v', '', '']],
                        columns: [{ type: 'hidden' }, { type: 'text', readOnly: true, align: 'left' }, { type: 'text', align: 'right' }],
                    },
                ],
            })[0];

            const cell = worksheet.records[0][0].element;

            worksheet.setMerge('A1', 3, 1);
            expect(worksheet.isReadOnly('A1')).to.equal(true);
            expect(cell.style.textAlign).to.equal('left');

            worksheet.hideColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(false);
            expect(cell.style.textAlign).to.equal('right');

            worksheet.showColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(true);
            expect(cell.style.textAlign).to.equal('left');
        });

        it('uses the options of the first visible column when deleteColumn and insertColumn move the range before hideColumn', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a', 'v', '', '']],
                        columns: [{ type: 'text' }, { type: 'hidden' }, { type: 'text', readOnly: true, align: 'left' }, { type: 'text', align: 'right' }],
                        mergeCells: { B1: [3, 1] },
                    },
                ],
            })[0];

            worksheet.deleteColumn(0);
            worksheet.hideColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(false);
            expect(worksheet.records[0][0].element.style.textAlign).to.equal('right');

            worksheet.showColumn(1);
            expect(worksheet.isReadOnly('A1')).to.equal(true);
            expect(worksheet.records[0][0].element.style.textAlign).to.equal('left');

            worksheet.insertColumn(1, 0, true);
            worksheet.hideColumn(2);
            expect(worksheet.isReadOnly('B1')).to.equal(false);
            expect(worksheet.records[0][1].element.style.textAlign).to.equal('right');
        });

        it('sends oncreatecell once for the first cell and keeps the changes of the handler', () => {
            const cells: HTMLTableCellElement[] = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '', 'x']],
                        columns: [{ type: 'hidden' }, { type: 'text' }, { type: 'text' }],
                        mergeCells: { A1: [2, 1] },
                    },
                ],
                oncreatecell: (worksheet, td, x, y) => {
                    if (x === 0 && y === 0) {
                        cells.push(td);
                        const badge = document.createElement('span');
                        badge.className = 'badge';
                        td.appendChild(badge);
                    }
                },
            })[0];

            expect(cells.length).to.equal(1);
            expect(worksheet.records[0][0].element.querySelector('.badge')).to.not.equal(null);
        });
    });
});

describe('Insert columns', () => {
    it('gives onbeforeinsertcolumn the column that insertColumn inserts without arguments', () => {
        let columns: Parameters<NonNullable<Jspreadsheet.SpreadsheetOptions['onbeforeinsertcolumn']>>[1] | undefined;
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            onbeforeinsertcolumn(instance, list) {
                columns = list;
            },
        })[0];

        worksheet.insertColumn();

        expect(columns!.length).to.eql(1);
        expect(columns![0].column).to.eql(2);
        expect(worksheet.getData()).to.eql([['a', 'b', '']]);
    });
});

describe('Drag a column header', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    // jsdom has no layout, so offsetX and offsetY are set: a press at the bottom edge of a header starts a drag
    const mouse = (element: Element, type: string) => {
        const event = new window.MouseEvent(type, { bubbles: true, buttons: 1 });
        Object.defineProperty(event, 'offsetX', { value: 0 });
        Object.defineProperty(event, 'offsetY', { value: 0 });
        element.dispatchEvent(event);
    };

    it('moves the column next to a header after column 9', async () => {
        const columns: Jspreadsheet.Column[] = [];
        for (let i = 0; i < 12; i++) {
            columns.push({ title: `C${i}` });
        }
        // Without column resize, the press at the edge is for a drag
        const worksheet = await load({ worksheets: [{ minDimensions: [12, 1], columns, columnResize: false }] });

        mouse(worksheet.headers[2], 'mousedown');
        // jsdom gives the header no width, so the pointer is on its right half
        mouse(worksheet.headers[10], 'mousemove');
        mouse(worksheet.headers[10], 'mouseup');

        expect(worksheet.getHeaders(true)).to.eql(['C0', 'C1', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C2', 'C11']);
    });

    it('keeps the index of the dragged column as a number', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 1], columnResize: false }] });

        mouse(worksheet.headers[2], 'mousedown');

        expect(worksheet.pointer.dragging!.column).to.equal(2);
        expect(worksheet.pointer.dragging!.destination).to.equal(2);

        mouse(worksheet.headers[2], 'mouseup');
    });

    it('keeps the index of the resized column as a number', async () => {
        // jsdom gives the header no width, so a press is at its right edge, which resizes the column
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 1] }] });

        mouse(worksheet.headers[2], 'mousedown');

        expect(worksheet.pointer.resizing!.column).to.equal(2);

        mouse(worksheet.headers[2], 'mouseup');
    });
});

describe('Delete columns with the keyboard', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    const mouse = (element: Element, type: string) => {
        const event = new window.MouseEvent(type, { bubbles: true, buttons: 1 });
        Object.defineProperty(event, 'offsetX', { value: 0 });
        Object.defineProperty(event, 'offsetY', { value: 0 });
        element.dispatchEvent(event);
    };

    let confirm: typeof global.confirm;
    beforeEach(() => {
        confirm = global.confirm;
    });
    afterEach(() => {
        global.confirm = confirm;
    });

    it('deletes the column whose header was pressed', async () => {
        const worksheet = await load({
            worksheets: [
                {
                    data: [
                        ['a', 'b', 'c'],
                        ['d', 'e', 'f'],
                    ],
                    columnResize: false,
                    columnDrag: false,
                },
            ],
        });

        const messages: (string | undefined)[] = [];
        global.confirm = (message) => {
            messages.push(message);
            return true;
        };

        mouse(worksheet.headers[1], 'mousedown');
        mouse(worksheet.headers[1], 'mouseup');
        document.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, key: 'Delete', keyCode: 46, which: 46 }));

        expect(messages).to.eql(['Are you sure to delete the selected columns?']);
        expect(worksheet.getData()).to.eql([
            ['a', 'c'],
            ['d', 'f'],
        ]);
    });
});

describe('Image columns', () => {
    it('shows a cell whose value is not text, and an image for image data', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [5, 'data:image/png;base64,iVBORw0KGgo='],
                        [true, 'a'],
                    ],
                    columns: [{ type: 'image' }, { type: 'image' }],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').querySelector('img')).to.equal(null);
        expect(worksheet.getCell('A2').querySelector('img')).to.equal(null);
        expect(worksheet.getCell('B1').querySelector('img')!.getAttribute('src')).to.equal('data:image/png;base64,iVBORw0KGgo=');
        expect(worksheet.getCell('B2').querySelector('img')).to.equal(null);
        expect(worksheet.getData()).to.eql([
            [5, 'data:image/png;base64,iVBORw0KGgo='],
            [true, 'a'],
        ]);
    });
});

describe('Calendar columns', () => {
    it('shows a cell whose value is a year, and sets a year as the value', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['2020'], ['2021-05-01']],
                    columns: [{ type: 'calendar' }],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1')).to.not.equal(undefined);
        expect(worksheet.getCell('A2').textContent).to.equal('2021-05-01');

        worksheet.setValueFromCoords(0, 1, '2022');

        expect(worksheet.getData()).to.eql([['2020'], ['2022']]);
    });
});

describe('Insert a column with nested headers', () => {
    it('widens the last nested header', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 2], nestedHeaders: [[{ title: 'All', colspan: 3 }]] }],
        })[0];

        worksheet.insertColumn(2);

        expect(worksheet.options.nestedHeaders![0][0].colspan).to.equal(5);
        const header = worksheet.thead.children[0].lastElementChild!;
        expect(header.getAttribute('colspan')).to.equal('5');
        expect(header.getAttribute('data-column')).to.equal('0,1,2,3,4');
        expect(worksheet.headers.map((header) => header.textContent)).to.eql(['A', 'B', 'C', 'D', 'E']);
        expect(worksheet.cols.map((col) => col.x)).to.eql([0, 1, 2, 3, 4]);
    });
});

describe('Insert columns', () => {
    it('gives the events the new columns, renders them and puts them before or after the column', () => {
        const events: [string, unknown][] = [];
        const rendered: [Jspreadsheet.CellValue | undefined, number, number][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
            onbeforeinsertcolumn: (instance, columns) => {
                events.push(['before', JSON.parse(JSON.stringify(columns))]);
            },
            oninsertcolumn: (instance, columns) => {
                events.push(['after', JSON.parse(JSON.stringify(columns))]);
            },
        })[0];

        worksheet.insertColumn(['x', 'y'], 0, true, [{ title: 'New', render: (td, value, x, y) => rendered.push([value, x, y]) }]);
        worksheet.insertColumn(2, 2, false);

        expect(worksheet.getData()).to.eql([
            ['x', 'a', 'b', '', ''],
            ['y', 'c', 'd', '', ''],
        ]);
        expect(rendered).to.eql([
            ['x', 0, 0],
            ['y', 0, 1],
        ]);
        expect(worksheet.headers.map((header) => header.textContent)).to.eql(['New', 'B', 'C', 'D', 'E']);
        expect(events).to.eql([
            ['before', [{ column: 0, options: { title: 'New' }, data: ['x', 'y'] }]],
            ['after', [{ column: 0, options: { title: 'New' }, data: ['x', 'y'] }]],
            [
                'before',
                [
                    { column: 3, options: {} },
                    { column: 4, options: {} },
                ],
            ],
            [
                'after',
                [
                    { column: 3, options: {} },
                    { column: 4, options: {} },
                ],
            ],
        ]);
        expect(worksheet.records.map((row) => row.map((cell) => cell.x))).to.eql([
            [0, 1, 2, 3, 4],
            [0, 1, 2, 3, 4],
        ]);
        expect(Array.from(worksheet.rows[0].element.children).slice(1)).to.eql(worksheet.records[0].map((cell) => cell.element));
    });
});

describe('A column of width 0', () => {
    it('keeps the width 0 of the options in the column and in getWidth', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], columns: [{ width: 80 }, { width: 0 }, {}, { width: '' }], defaultColWidth: 50 }],
        })[0];

        expect(worksheet.cols[1].colElement.getAttribute('width')).to.equal('0');
        expect(worksheet.getWidth(1)).to.equal(0);
        expect(worksheet.getWidth()).to.eql([80, 0, 50, 50]);
    });

    it('keeps a default width of 0', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 1], defaultColWidth: 0 }] })[0];

        expect(worksheet.cols[0].colElement.getAttribute('width')).to.equal('0');
        expect(worksheet.getWidth()).to.eql([0, 0]);
    });

    it('sets a width of 0 and gives back a width of 0 with undo', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 1], columns: [{ width: 80 }, { width: 0 }, { width: 70 }] }],
        })[0];

        worksheet.setWidth(1, 50);
        worksheet.undo();
        expect(worksheet.getWidth(1)).to.equal(0);
        expect(worksheet.getWidth()).to.eql([80, 0, 70]);

        worksheet.setWidth([0, 1], [50, 40]);
        worksheet.undo();
        expect(worksheet.getWidth()).to.eql([80, 0, 70]);

        worksheet.setWidth(2, 0);
        expect(worksheet.getWidth(2)).to.equal(0);
        worksheet.undo();
        expect(worksheet.getWidth(2)).to.equal(70);

        // NaN is no width, so the columns keep their widths
        worksheet.setWidth(2, NaN);
        worksheet.setWidth([0, 1], [NaN, NaN]);
        expect(worksheet.getWidth()).to.eql([80, 0, 70]);
    });
});

describe('The keyboard and a column of width 0', () => {
    const range = (worksheet: Worksheet) => worksheet.selection.range!.map(Number);

    it('does not make a cell in a column of width 0 the first tab stop', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [5, 2], columns: [{ width: 0 }, {}, { width: 0 }, {}, { width: 0 }] }],
        })[0];

        expect(Array.from(worksheet.table.querySelectorAll('[tabindex="0"]'))).to.eql([worksheet.records[0][1].element]);
    });

    it('moves over a column of width 0 with the arrow keys, Ctrl, Shift, Home and End', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [5, 2], columns: [{ width: 0 }, {}, { width: 0 }, {}, { width: 0 }] }],
        })[0];

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        worksheet.right();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
        worksheet.right();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
        worksheet.left();
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);
        worksheet.left();
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);

        worksheet.right(false, true);
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
        worksheet.left(false, true);
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);

        worksheet.right(true);
        expect(range(worksheet)).to.eql([1, 0, 3, 0]);
        worksheet.right(true);
        expect(range(worksheet)).to.eql([1, 0, 3, 0]);

        worksheet.updateSelectionFromCoords(3, 0, 3, 0);
        worksheet.first();
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);
        worksheet.last();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
    });

    it('pastes into a column of width 0, so that paste fills the cells that copy reads', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], columns: [{}, { width: 0 }, {}, { width: 0 }] }],
        })[0];

        worksheet.paste(0, 0, '1\t2\t3\t4');

        expect(worksheet.getData()[0]).to.eql(['1', '2', '3', '4']);
    });

    it('pastes into a column of width 0 in the row source mode', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [4, 1],
                    columns: [{}, { width: 0 }, {}, { width: 0 }],
                    rowSource: { count: 3, row: () => ['', '', '', ''] },
                },
            ],
        })[0];

        worksheet.paste(0, 0, '1\t2\t3\t4');

        expect([0, 1, 2, 3].map((x) => worksheet.getValueFromCoords(x, 0))).to.eql(['1', '2', '3', '4']);
    });

    it('stops on a merged range that starts in a column of width 0 and shows in the next column', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 1], columns: [{}, { width: 0 }, {}, {}], mergeCells: { B1: [2, 1] } }],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.right();
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);
        worksheet.right();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
    });

    it('moves over a merged range whose columns all have the width 0', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 1], columns: [{}, { width: 0 }, { width: 0 }, {}], mergeCells: { B1: [2, 1] } }],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.right();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
        worksheet.left();
        expect(range(worksheet)).to.eql([0, 0, 0, 0]);
    });

    it('stops on a merged range that starts in a column of width 0 in the row source mode', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [4, 1],
                    columns: [{}, { width: 0 }, {}, {}],
                    mergeCells: { B1: [2, 1] },
                    rowSource: { count: 10, row: (y) => [`a${y}`, `b${y}`, `c${y}`, `d${y}`] },
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.right();
        expect(range(worksheet)).to.eql([1, 0, 1, 0]);
        worksheet.right();
        expect(range(worksheet)).to.eql([3, 0, 3, 0]);
    });

    it('moves over a column of width 0 in the row source mode', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [4, 1],
                    columns: [{}, { width: 0 }, {}, { width: 0 }],
                    rowSource: { count: 10, row: (y) => [`a${y}`, `b${y}`, `c${y}`, `d${y}`] },
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.right();
        expect(range(worksheet)).to.eql([2, 0, 2, 0]);
        worksheet.right(false, true);
        expect(range(worksheet)).to.eql([2, 0, 2, 0]);
        worksheet.left();
        expect(range(worksheet)).to.eql([0, 0, 0, 0]);
    });
});

describe('Set the widths of columns with an array', () => {
    it('setWidth keeps the width of a column that the array of widths has no width for (issue #47)', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 2], defaultColWidth: 80 }] })[0];

        worksheet.setWidth([0, 1, 2], [100, 120]);

        expect(worksheet.getWidth(0)).to.equal(100);
        expect(worksheet.getWidth(1)).to.equal(120);
        expect(worksheet.getWidth(2)).to.equal(80);
        expect(worksheet.options.columns[2].width).to.not.be.an('array');

        worksheet.setWidth([0, 1], [90]);

        expect(worksheet.getWidth(0)).to.equal(90);
        expect(worksheet.getWidth(1)).to.equal(120);
        expect(worksheet.options.columns[1].width).to.equal(120);

        worksheet.undo();

        expect(worksheet.getWidth(0)).to.equal(100);
        expect(worksheet.getWidth(1)).to.equal(120);
    });
});

describe('The first cell of a merged range over a hidden column (issue #20)', () => {
    it('gives a formula its number with the decimal point of the first visible column', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1,5', '', '=A1+1']],
                    columns: [{ type: 'hidden' }, { type: 'text', decimal: ',' }, { type: 'text' }],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        })[0];

        expect(worksheet.getValueFromCoords(2, 0, true)).to.equal('2.5');
    });

    it('reads a pasted date with the format of the first visible column', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['', '', '']],
                    columns: [{ type: 'hidden' }, { type: 'calendar', options: { format: 'DD/MM/YYYY' } }, { type: 'text' }],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        })[0];

        worksheet.paste(0, 0, '25/12/2024');

        expect(worksheet.getValueFromCoords(0, 0)).to.match(/^2024-12-25/);
    });

    it('moves the references of a formula that the fill handle copies when the first visible column is a text column', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['=C2', '', ''],
                        ['', '', '7'],
                    ],
                    columns: [{ type: 'text' }, { type: 'hidden' }, { type: 'text' }],
                    mergeCells: { B1: [2, 1] },
                },
            ],
        })[0];

        // Fill the merged range B1:C1 from A1
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.pointer.fill(worksheet.records[0][1].element, worksheet.records[0][1].element);

        expect(worksheet.getValueFromCoords(1, 0)).to.equal('=D2');
    });

    it('fills down from a merged range over a hidden first column with a double click on the corner (issue #127)', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '', ''],
                        ['', '', ''],
                    ],
                    columns: [{ type: 'hidden' }, { type: 'text' }, { type: 'text' }],
                    mergeCells: { A1: [2, 1], A2: [2, 1] },
                },
            ],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.pointer.fillDown();

        expect(worksheet.getValueFromCoords(0, 1)).to.equal(2);
    });

    it('wraps a long text when hideColumn renders the first cell with the options of another column', () => {
        const long = 'a'.repeat(250);
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1', '', 'x']],
                    columns: [{ type: 'hidden' }, { type: 'dropdown', source: [{ id: '1', name: long }] }, { type: 'text' }],
                    mergeCells: { A1: [3, 1] },
                },
            ],
        })[0];

        const cell = worksheet.records[0][0].element;
        expect(cell.style.whiteSpace).to.equal('pre-wrap');

        worksheet.hideColumn(1);
        expect(cell.textContent).to.equal('1');
        expect(cell.style.whiteSpace).to.equal('');

        worksheet.showColumn(1);
        expect(cell.style.whiteSpace).to.equal('pre-wrap');
    });

    it('sends oncreatecell when hideColumn and showColumn render the first cell again', () => {
        const calls: [HTMLTableCellElement, Jspreadsheet.CellValue][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1', '', 'x']],
                    columns: [{ type: 'hidden' }, { type: 'text' }, { type: 'text' }],
                    mergeCells: { A1: [3, 1] },
                },
            ],
            oncreatecell: (instance, td, x, y, value) => {
                if (x === 0 && y === 0) {
                    calls.push([td, value]);
                }
            },
        })[0];

        calls.length = 0;
        worksheet.hideColumn(1);
        worksheet.showColumn(1);

        const cell = worksheet.records[0][0].element;
        expect(calls).to.eql([
            [cell, '1'],
            [cell, '1'],
        ]);
    });
});

describe('The worksheet option wrapLongText', () => {
    const long = 'a'.repeat(250);

    it('wraps a long text by default', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [[long, 'b']] }] })[0];

        expect(worksheet.records[0][0].element.style.whiteSpace).to.equal('pre-wrap');
    });

    it('does not wrap a long text when it is false, also after setValue and a render of a merged cell', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [long, 'b', 'c', 'd'],
                        ['1', '', 'x', ''],
                    ],
                    columns: [{ type: 'hidden' }, { type: 'dropdown', source: [{ id: '1', name: long }] }, { type: 'text', wordWrap: true }, { type: 'text' }],
                    mergeCells: { A2: [3, 1] },
                    wrapLongText: false,
                },
            ],
        })[0];

        expect(worksheet.records[0][0].element.style.whiteSpace).to.equal('');
        worksheet.setValueFromCoords(3, 0, long);
        expect(worksheet.records[0][3].element.style.whiteSpace).to.equal('');

        // The merged range shows the text of column 2, which wraps, and then the long text of column 1, which does not
        const merged = worksheet.records[1][0].element;
        worksheet.hideColumn(1);
        expect(merged.style.whiteSpace).to.equal('pre-wrap');
        worksheet.showColumn(1);
        expect(merged.style.whiteSpace).to.equal('');
    });

    it('keeps the wordWrap options of the worksheet and of a column when it is false', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']], columns: [{ wordWrap: true }, {}], wrapLongText: false }],
        })[0];
        expect(worksheet.records[0][0].element.style.whiteSpace).to.equal('pre-wrap');
        expect(worksheet.records[0][1].element.style.whiteSpace).to.equal('');

        const all = jspreadsheet(document.createElement('div'), { worksheets: [{ data: [['a']], wordWrap: true, wrapLongText: false }] })[0];
        expect(all.records[0][0].element.style.whiteSpace).to.equal('pre-wrap');
    });
});
