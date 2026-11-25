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
        expect(worksheet.selectedCell).to.eql([1, 1, 1, 1]);

        press(worksheet.records[2][2].element, { shiftKey: true });
        expect(worksheet.selectedCell).to.eql([1, 1, 2, 2]);
    });

    it('extends the selection to column 0 and row 0 when the mouse moves over them', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        press(worksheet.records[1][1].element);
        worksheet.records[0][0].element.dispatchEvent(new window.MouseEvent('mouseover', { bubbles: true, buttons: 1 }));
        expect(worksheet.selectedCell).to.eql([1, 1, 0, 0]);
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
        worksheet.paste(worksheet.selectedCell[0], worksheet.selectedCell[1], 'a');

        expect(coords).to.eql([2, 1]);
        expect(worksheet.getValue('C2')).to.equal('a');
    });
});
