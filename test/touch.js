const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Touch', () => {
    it('opens no editor after a long touch when the worksheet is not editable', async () => {
        const events = [];
        const worksheet = await new Promise((resolve) => {
            const instance = jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3], editable: false }],
                oneditionstart: () => events.push('start'),
                onload: () => resolve(instance[0]),
            });
        });

        worksheet.records[1][1].element.dispatchEvent(new window.Event('touchstart', { bubbles: true }));

        // A long touch opens the editor after 500 ms
        await new Promise((resolve) => setTimeout(resolve, 600));

        expect(worksheet.edition).to.not.be.ok;
        expect(events).to.eql([]);
        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });
});
