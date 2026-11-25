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

    it('keeps the scroll position when ArrowLeft selects column A at a small scroll and the index column is hidden', () => {
        const worksheet = scrolledWorksheet(40);
        worksheet.hideIndex();

        // Without the index column, the columns are frozen at each scroll, so column A stays at the left of the content
        worksheet.records[1][0].element.getBoundingClientRect = rect(0, 30, 100, 25);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.left();

        expect(worksheet.selectedCell.map(Number)).to.eql([0, 1, 0, 1]);
        expect(worksheet.content.scrollLeft).to.equal(40);
    });
});

describe('Freeze columns with a hidden index column', () => {
    const scroll = function (worksheet, scrollLeft) {
        // The scroll handler compares scrollLeft with the value of the last scroll, so start each test at 0
        worksheet.content.scrollLeft = 0;
        worksheet.content.dispatchEvent(new window.Event('scroll'));
        worksheet.content.scrollLeft = scrollLeft;
        worksheet.content.dispatchEvent(new window.Event('scroll'));
    };

    it('puts the frozen cells at the left border of the content when hideIndex hides the index column', () => {
        const worksheet = scrolledWorksheet(0);
        worksheet.hideIndex();

        scroll(worksheet, 900);

        expect(worksheet.records[1][0].element.classList.contains('jss_freezed')).to.equal(true);
        expect(worksheet.records[1][0].element.style.left).to.equal('899px');
        expect(worksheet.records[1][1].element.style.left).to.equal('899px');
    });

    it('freezes the columns at a small scroll when hideIndex hides the index column', () => {
        const worksheet = scrolledWorksheet(0);
        worksheet.hideIndex();

        scroll(worksheet, 20);

        expect(worksheet.headers[0].classList.contains('jss_freezed')).to.equal(true);
        expect(worksheet.records[1][0].element.classList.contains('jss_freezed')).to.equal(true);
        expect(worksheet.records[1][0].element.style.left).to.equal('19px');
    });

    it('freezes the columns only when the scroll is more than the width of a visible index column', () => {
        const worksheet = scrolledWorksheet(0);

        scroll(worksheet, 20);
        expect(worksheet.records[1][0].element.classList.contains('jss_freezed')).to.equal(false);

        scroll(worksheet, 900);
        expect(worksheet.records[1][0].element.classList.contains('jss_freezed')).to.equal(true);
        expect(worksheet.records[1][0].element.style.left).to.equal('849px');
    });

    it('puts the frozen cells at the left border of the content when hideIndex is called after the scroll', () => {
        const worksheet = scrolledWorksheet(0);

        scroll(worksheet, 900);
        worksheet.hideIndex();

        expect(worksheet.records[1][0].element.style.left).to.equal('899px');
        expect(worksheet.records[1][1].element.style.left).to.equal('899px');
    });

    it('puts the frozen cells to the right of the index column when showIndex is called after the scroll', () => {
        const worksheet = scrolledWorksheet(0);
        worksheet.hideIndex();

        scroll(worksheet, 900);
        worksheet.showIndex();

        expect(worksheet.records[1][0].element.style.left).to.equal('849px');
        expect(worksheet.records[1][1].element.style.left).to.equal('849px');
    });

    it('unfreezes the columns when showIndex is called at a small scroll', () => {
        const worksheet = scrolledWorksheet(0);
        worksheet.hideIndex();

        scroll(worksheet, 20);
        worksheet.showIndex();

        expect(worksheet.headers[0].classList.contains('jss_freezed')).to.equal(false);
        expect(worksheet.records[1][0].element.classList.contains('jss_freezed')).to.equal(false);
        expect(worksheet.records[1][0].element.style.left).to.equal('');
    });

    it('hides and shows the index column when the worksheet has fewer columns than freezeColumns', () => {
        const worksheet = scrolledWorksheet(0);
        worksheet.deleteColumn(0, 11);

        expect(() => worksheet.hideIndex()).to.not.throw();
        expect(() => worksheet.showIndex()).to.not.throw();

        // The frozen columns are placed only when the scroll is more than the width of the index column
        scroll(worksheet, 900);
        expect(() => worksheet.hideIndex()).to.not.throw();
        expect(() => worksheet.showIndex()).to.not.throw();
    });
});
