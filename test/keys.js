const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

// The spreadsheet adds its event handlers after a promise
const load = (options) =>
    new Promise((resolve) => {
        const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
    });

describe('Keys', () => {
    const ctrlKey = function (target, key, keyCode) {
        target.dispatchEvent(
            new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: key, keyCode: keyCode, which: keyCode, ctrlKey: true, metaKey: true })
        );
    };

    it('copies the selection with Ctrl+X when the worksheet is read-only', async () => {
        const copied = [];
        const worksheet = await load({
            worksheets: [{ data: [['a', 'b']], editable: false }],
            oncopy: function (instance, coords, text, isCut) {
                copied.push(text);
            },
        });

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);
        ctrlKey(worksheet.records[0][0].element, 'x', 88);

        expect(copied).to.eql(['a\tb']);
        expect(worksheet.getData()).to.eql([['a', 'b']]);
    });

    it('moves the selection to the first cell of a merged range, not to a cell that the range hides', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 6], mergeCells: { A1: [2, 1], C2: [1, 3] } }],
        })[0];

        // Up from C5 enters the range C2:C4 at C4
        worksheet.updateSelectionFromCoords(2, 4, 2, 4);
        worksheet.up();
        expect(worksheet.selection.range.map(Number)).to.eql([2, 1, 2, 1]);

        // Left from C1 enters the range A1:B1 at B1
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);
        worksheet.left();
        expect(worksheet.selection.range.map(Number)).to.eql([0, 0, 0, 0]);

        // The next key starts from the first cell of the range
        worksheet.down();
        expect(worksheet.selection.range.map(Number)).to.eql([0, 1, 0, 1]);
    });

    it('moves the selection to the first cell of a merged range after End', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], mergeCells: { C1: [2, 1] } }],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.last();
        expect(worksheet.selection.range.map(Number)).to.eql([2, 0, 2, 0]);
    });

    it('moves the selection to the first cell of a merged range after Right, Down and Home', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 6], mergeCells: { A1: [1, 3], C2: [1, 3], A5: [2, 1] } }],
        })[0];

        // Right from B3 enters the range C2:C4 at C3
        worksheet.updateSelectionFromCoords(1, 2, 1, 2);
        worksheet.right();
        expect(worksheet.selection.range.map(Number)).to.eql([2, 1, 2, 1]);

        // Down from B4 enters the range A5:B5 at B5
        worksheet.updateSelectionFromCoords(1, 3, 1, 3);
        worksheet.down();
        expect(worksheet.selection.range.map(Number)).to.eql([0, 4, 0, 4]);

        // Home from D3 enters the range A1:A3 at A3
        worksheet.updateSelectionFromCoords(3, 2, 3, 2);
        worksheet.first();
        expect(worksheet.selection.range.map(Number)).to.eql([0, 0, 0, 0]);
    });
});

describe('Toggle a checkbox with the space key', () => {
    it('gives the change events the column and row as numbers', async () => {
        const calls = [];
        const worksheet = await load({
            worksheets: [{ data: [['a', false]], columns: [{ type: 'text' }, { type: 'checkbox' }] }],
            onbeforechange: (instance, cell, x, y) => {
                calls.push(['onbeforechange', x, y]);
            },
            onafterchanges: (instance, changes) => {
                calls.push(['onafterchanges', changes[0].x, changes[0].y]);
            },
        });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        document.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, key: ' ', keyCode: 32, which: 32 }));

        expect(worksheet.getValueFromCoords(1, 0)).to.equal(true);
        expect(calls).to.eql([
            ['onbeforechange', 1, 0],
            ['onafterchanges', 1, 0],
        ]);
    });
});

describe('Moving the selection without a selection', () => {
    it('does nothing in up, down, left, right, first and last', () => {
        const events = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                },
            ],
            onselection: () => events.push('onselection'),
        })[0];

        for (const method of ['up', 'down', 'left', 'right', 'first', 'last']) {
            worksheet[method]();
            worksheet[method](true, true);
        }

        expect(worksheet.getSelection()).to.equal(null);
        expect(events).to.deep.equal([]);
    });

    it('does nothing after resetSelection', () => {
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
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.resetSelection();

        worksheet.down();
        worksheet.right();

        expect(worksheet.getSelection()).to.equal(null);
    });
});
