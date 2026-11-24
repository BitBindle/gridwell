const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const rect = function (left, top, width, height) {
    return () => ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top });
};

const scrolledWorksheet = function (scrollLeft) {
    const data = [];
    for (let j = 0; j < 5; j++) {
        data.push([]);
        for (let i = 0; i < 12; i++) {
            data[j].push(i + ':' + j);
        }
    }

    const worksheet = jspreadsheet(root, {
        worksheets: [{ data, freezeColumns: 2, defaultColWidth: 300, columns: [{ width: 100 }, { width: 100 }], tableOverflow: true }],
    })[0];

    // The worksheet is 600 pixels wide. The frozen columns are 200 pixels wide.
    Object.defineProperty(worksheet.content, 'scrollLeft', { value: scrollLeft, writable: true });
    worksheet.content.getBoundingClientRect = rect(0, 0, 600, 300);

    return worksheet;
};

describe('Freeze columns', () => {
    it('ArrowLeft puts the active cell to the right of the frozen columns, not under them', () => {
        const worksheet = scrolledWorksheet(900);

        // Column H starts at the left border of the content, under the frozen columns
        worksheet.records[1][7].element.getBoundingClientRect = rect(0, 30, 300, 25);

        worksheet.updateSelectionFromCoords(8, 1, 8, 1);
        worksheet.left();

        expect(worksheet.selectedCell.map(Number)).to.eql([7, 1, 7, 1]);
        expect(worksheet.content.scrollLeft).to.equal(900 - 200 - 1);
    });

    it('keeps the scroll position when ArrowLeft selects a cell of a frozen column', () => {
        const worksheet = scrolledWorksheet(900);

        // The frozen column B stays at the left of the content
        worksheet.records[1][1].element.getBoundingClientRect = rect(100, 30, 100, 25);

        worksheet.updateSelectionFromCoords(2, 1, 2, 1);
        worksheet.left();

        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
        expect(worksheet.content.scrollLeft).to.equal(900);
    });

    it('ArrowRight from a frozen column puts the active cell to the right of the frozen columns', () => {
        const worksheet = scrolledWorksheet(900);

        // Column C starts 200 pixels from the left of the table, so it is far to the left of the content
        worksheet.records[1][2].element.getBoundingClientRect = rect(-700, 30, 300, 25);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.right();

        expect(worksheet.selectedCell.map(Number)).to.eql([2, 1, 2, 1]);
        expect(worksheet.content.scrollLeft).to.equal(0);
    });

    it('scrolls a frozen column into view when the scroll is too small to freeze it', () => {
        const worksheet = scrolledWorksheet(40);
        worksheet.hideIndex();

        // At scrollLeft 40 the columns are not frozen, and 40 pixels of column A are to the left of the content
        worksheet.records[1][0].element.getBoundingClientRect = rect(-40, 30, 100, 25);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.left();

        expect(worksheet.selectedCell.map(Number)).to.eql([0, 1, 0, 1]);
        expect(worksheet.content.scrollLeft).to.equal(0);
    });
});
