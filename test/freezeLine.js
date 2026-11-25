const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const frozenWorksheet = function (columns) {
    const worksheet = jspreadsheet(root, {
        worksheets: [{ minDimensions: [12, 3], freezeColumns: 3, columns, tableOverflow: true }],
    })[0];

    // Each worksheet shares the last scroll position, so start from 0
    for (const scrollLeft of [0, 300]) {
        worksheet.content.scrollLeft = scrollLeft;
        worksheet.content.dispatchEvent(new Event('scroll'));
    }

    return worksheet;
};

const lastFrozen = function (worksheet) {
    const cells = [worksheet.headers].concat(worksheet.records.map((row) => row.map((cell) => cell.element)));

    return cells.map((row) => row.filter((cell) => cell.classList.contains('jss_freezed_last')).map((cell) => Number(cell.getAttribute('data-x'))));
};

describe('Freeze columns line', () => {
    it('marks the header and the cells of the last frozen column when the worksheet is scrolled', () => {
        const worksheet = frozenWorksheet();

        expect(worksheet.records[0][2].element.classList.contains('jss_freezed')).to.equal(true);
        expect(lastFrozen(worksheet)).to.eql([[2], [2], [2], [2]]);
    });

    it('marks the last frozen column that is not hidden', () => {
        const worksheet = frozenWorksheet([{}, {}, { type: 'hidden' }]);

        expect(lastFrozen(worksheet)).to.eql([[1], [1], [1], [1]]);
    });

    it('moves the mark when the last frozen column is hidden and shown again', () => {
        const worksheet = frozenWorksheet();

        worksheet.hideColumn(2);
        expect(lastFrozen(worksheet)).to.eql([[1], [1], [1], [1]]);

        worksheet.showColumn(2);
        expect(lastFrozen(worksheet)).to.eql([[2], [2], [2], [2]]);
    });

    it('marks the last column when the worksheet has fewer columns than freezeColumns', () => {
        const worksheet = frozenWorksheet();

        worksheet.deleteColumn(1, 11);
        worksheet.content.scrollLeft = 400;
        worksheet.content.dispatchEvent(new Event('scroll'));

        expect(lastFrozen(worksheet)).to.eql([[0], [0], [0], [0]]);
    });

    it('removes the mark when the worksheet is scrolled back', () => {
        const worksheet = frozenWorksheet();

        worksheet.content.scrollLeft = 0;
        worksheet.content.dispatchEvent(new Event('scroll'));

        expect(lastFrozen(worksheet)).to.eql([[], [], [], []]);
        expect(worksheet.records[0][2].element.classList.contains('jss_freezed')).to.equal(false);
    });
});
