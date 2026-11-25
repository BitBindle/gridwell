const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

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
    const load = (options) =>
        new Promise((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });
    const press = (element, options) => element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, buttons: 1, ...options }));

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
        const coords = [];
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            onbeforepaste: (instance, data, x, y) => {
                coords.push(x, y);
            },
        });

        press(worksheet.records[1][2].element);
        worksheet.paste(worksheet.selection.range[0], worksheet.selection.range[1], 'a');

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
    const load = (options) =>
        new Promise((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });
    const mouse = (element, type, options) => element.dispatchEvent(new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...options }));

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
