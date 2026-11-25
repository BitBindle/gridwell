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
