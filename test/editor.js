const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const rows = function (count) {
    const data = [];
    for (let j = 0; j < count; j++) {
        data.push(['' + j, '']);
    }
    return data;
};

describe('Editor', () => {
    it('keeps the text editor open when the grid scrolls', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[1][0].element);
        worksheet.content.dispatchEvent(new window.Event('scroll'));

        expect(worksheet.edition).to.not.equal(null);
    });

    it('closes a dropdown editor when the grid scrolls', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: rows(300),
                    columns: [{ type: 'dropdown', source: ['0', '1', '2'] }, {}],
                    lazyLoading: true,
                    tableOverflow: true,
                },
            ],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[1][0].element);
        worksheet.content.dispatchEvent(new window.Event('scroll'));

        expect(worksheet.edition).to.equal(null);
    });

    it('saves and closes the text editor when lazy loading removes its row', async () => {
        // The grid loads rows 100 ms after a scroll, one load at a time. Let the loads of the tests before end.
        await new Promise((resolve) => setTimeout(resolve, 150));

        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[0][1].element);
        worksheet.records[0][1].element.querySelector('input, textarea').value = 'typed';

        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
        await new Promise((resolve) => setTimeout(resolve, 150));

        expect(worksheet.tbody.contains(worksheet.records[0][1].element)).to.equal(false);
        expect(worksheet.edition).to.equal(null);
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('typed');
    });
});
