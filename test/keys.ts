import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// The spreadsheet adds its event handlers after a promise
const load = (options: Jspreadsheet.SpreadsheetOptions) =>
    new Promise<Worksheet>((resolve) => {
        const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
    });

describe('Keys', () => {
    function ctrlKey(target: Element, key: string, keyCode: number) {
        target.dispatchEvent(
            new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, keyCode, which: keyCode, ctrlKey: true, metaKey: true })
        );
    }

    it('copies the selection with Ctrl+X when the worksheet is read-only', async () => {
        const copied: string[] = [];
        const worksheet = await load({
            worksheets: [{ data: [['a', 'b']], editable: false }],
            oncopy(instance, coords, text, isCut) {
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
        expect(worksheet.selection.range!.map(Number)).to.eql([2, 1, 2, 1]);

        // Left from C1 enters the range A1:B1 at B1
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);
        worksheet.left();
        expect(worksheet.selection.range!.map(Number)).to.eql([0, 0, 0, 0]);

        // The next key starts from the first cell of the range
        worksheet.down();
        expect(worksheet.selection.range!.map(Number)).to.eql([0, 1, 0, 1]);
    });

    it('moves the selection to the first cell of a merged range after End', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], mergeCells: { C1: [2, 1] } }],
        })[0];

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.last();
        expect(worksheet.selection.range!.map(Number)).to.eql([2, 0, 2, 0]);
    });

    it('moves the selection to the first cell of a merged range after Right, Down and Home', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 6], mergeCells: { A1: [1, 3], C2: [1, 3], A5: [2, 1] } }],
        })[0];

        // Right from B3 enters the range C2:C4 at C3
        worksheet.updateSelectionFromCoords(1, 2, 1, 2);
        worksheet.right();
        expect(worksheet.selection.range!.map(Number)).to.eql([2, 1, 2, 1]);

        // Down from B4 enters the range A5:B5 at B5
        worksheet.updateSelectionFromCoords(1, 3, 1, 3);
        worksheet.down();
        expect(worksheet.selection.range!.map(Number)).to.eql([0, 4, 0, 4]);

        // Home from D3 enters the range A1:A3 at A3
        worksheet.updateSelectionFromCoords(3, 2, 3, 2);
        worksheet.first();
        expect(worksheet.selection.range!.map(Number)).to.eql([0, 0, 0, 0]);
    });
});

describe('Toggle a checkbox with the space key', () => {
    it('gives the change events the column and row as numbers', async () => {
        const calls: unknown[][] = [];
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
        const events: string[] = [];
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

        for (const method of ['up', 'down', 'left', 'right', 'first', 'last'] as const) {
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

describe('Keys on the grid', () => {
    function keydown(target: Element, key: string, keyCode: number, options: KeyboardEventInit = {}) {
        const event = new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, keyCode, which: keyCode, ...options });
        target.dispatchEvent(event);
        return event;
    }

    it('adds a row with Enter on the last row, and moves the selection with Tab and Shift+Tab', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b', 'c']] }] });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        expect(keydown(worksheet.records[0][1].element, 'Tab', 9).defaultPrevented).to.equal(true);

        expect(worksheet.getSelection()).to.eql([2, 0, 2, 0]);

        keydown(worksheet.records[0][2].element, 'Enter', 13);

        expect(worksheet.rows.length).to.equal(2);
        expect(worksheet.getSelection()).to.eql([2, 1, 2, 1]);

        keydown(worksheet.records[1][2].element, 'Enter', 13, { shiftKey: true });
        expect(keydown(worksheet.records[0][2].element, 'Tab', 9, { shiftKey: true }).defaultPrevented).to.equal(true);

        expect(worksheet.getSelection()).to.eql([1, 0, 1, 0]);
        expect(worksheet.rows.length).to.equal(2);
        expect(worksheet.headers.length).to.equal(3);
    });

    it('leaves the default action of Tab on the last column and of Shift+Tab on the first column, so that the focus leaves the grid (issue #28)', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']] }] });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        expect(keydown(worksheet.records[0][1].element, 'Tab', 9).defaultPrevented).to.equal(false);

        // No column is added, and the selection stays
        expect(worksheet.headers.length).to.equal(2);
        expect(worksheet.getSelection()).to.eql([1, 0, 1, 0]);

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        expect(keydown(worksheet.records[0][0].element, 'Tab', 9, { shiftKey: true }).defaultPrevented).to.equal(false);
        expect(worksheet.getSelection()).to.eql([0, 0, 0, 0]);
    });

    it('clears the selected cells with Delete', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b', 'c']] }] });

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);
        keydown(worksheet.records[0][0].element, 'Delete', 46);

        expect(worksheet.getData()).to.eql([['', '', 'c']]);
    });

    it('selects all with Ctrl+A, and undoes and redoes with Ctrl+Z and Ctrl+Y', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']] }] });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        keydown(worksheet.records[0][0].element, 'a', 65, { ctrlKey: true });

        // @ts-expect-error -- the test passes a second argument, which getSelected ignores
        expect(worksheet.getSelected(false, true).length).to.equal(2);

        worksheet.setValue('A1', 'x');
        keydown(worksheet.records[0][0].element, 'z', 90, { ctrlKey: true });

        expect(worksheet.getValue('A1')).to.equal('a');

        keydown(worksheet.records[0][0].element, 'y', 89, { ctrlKey: true });

        expect(worksheet.getValue('A1')).to.equal('x');
    });

    it('toggles a checkbox with Space and opens the editor with a character', async () => {
        const worksheet = await load({ worksheets: [{ data: [[false, 'a']], columns: [{ type: 'checkbox' }, {}] }] });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        keydown(worksheet.records[0][0].element, ' ', 32);

        expect(worksheet.getValue('A1')).to.equal(true);
        expect(worksheet.editor.isOpen()).to.equal(false);

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        keydown(worksheet.records[0][1].element, 'q', 81);

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(worksheet.editor.cell).to.equal(worksheet.records[0][1].element);
    });

    it('does not edit with the keys when the worksheet is not editable', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']], editable: false }] });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        keydown(worksheet.records[0][0].element, 'Delete', 46);
        keydown(worksheet.records[0][0].element, 'q', 81);

        expect(worksheet.getData()).to.eql([['a', 'b']]);
        expect(worksheet.editor.isOpen()).to.equal(false);
    });
});
