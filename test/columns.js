const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Use the columns method', () => {
    it('insertColumn and column is inserted in the position 0', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(9);

        // Insert [3, 3] column in the first column
        instance[0].insertColumn([3, 3], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(3);

        // Insert [2, 2] column in the first column
        instance[0].insertColumn([2, 2], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(2);
    });

    it('insertColumn and column is inserted in the position 1', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[2].innerHTML).to.include(9);

        // Insert [3, 3] column in the second column
        instance[0].insertColumn([3, 3], 0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(3);

        // Insert [2, 2] column in the second column
        instance[0].insertColumn([2, 2], 0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(2);
    });

    it('deleteColumn and column is removed in the given index', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 9, 9, 9],
                        [3, 6, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(3);
        expect(rows[0].children[2].innerHTML).to.include(6);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(6);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);
    });

    it('deleteColumn and multiple column are removed starting from the given index', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 12, 15, 18],
                        [3, 6, 9, 12, 15, 18],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(3);
        expect(rows[0].children[2].innerHTML).to.include(6);
        expect(rows[0].children[3].innerHTML).to.include(9);

        // Delete first two columns
        instance[0].deleteColumn(0, 2);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(12);
        expect(rows[0].children[3].innerHTML).to.include(15);

        // Delete first two columns
        instance[0].deleteColumn(0, 2);

        expect(rows[0].children[1].innerHTML).to.include(15);
        expect(rows[0].children[2].innerHTML).to.include(18);
    });

    it('hideColumn and showColumn', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [3, 6, 9, 12, 15, 18],
                        [3, 6, 9, 12, 15, 18],
                    ],
                },
            ],
        });

        let table = root.querySelector('thead');
        let headers = table.children[0].children;

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).not.to.include('none');

        // Hides the first column 'A'
        instance[0].hideColumn(0);

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).to.include('none');
        expect(headers[2].innerHTML).to.include('B');
        expect(window.getComputedStyle(headers[2]).display).not.to.include('none');

        // Shows the column that was hidden
        instance[0].showColumn(0);

        expect(headers[1].innerHTML).to.include('A');
        expect(window.getComputedStyle(headers[1]).display).not.to.include('none');
        expect(headers[2].innerHTML).to.include('B');
        expect(window.getComputedStyle(headers[2]).display).not.to.include('none');
    });

    it('insertColumn history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [9, 9, 9, 9, 9, 9],
                        [9, 9, 9, 9, 9, 9],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(9);

        // Insert [3, 3] column in the first column
        instance[0].insertColumn([3, 3], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(3);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(9);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(3);
    });

    it('deleteColumn history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3, 4],
                        [5, 6, 7, 8],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        // Delete first column
        instance[0].deleteColumn(0);

        expect(rows[0].children[1].innerHTML).to.include(2);
        expect(rows[0].children[2].innerHTML).to.include(3);
        expect(rows[0].children[3].innerHTML).to.include(4);

        instance[0].undo(0);

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].redo(0);

        expect(rows[0].children[1].innerHTML).to.include(2);
        expect(rows[0].children[2].innerHTML).to.include(3);
        expect(rows[0].children[3].innerHTML).to.include(4);
    });

    it('hideColumn keeps a merged range visible when it covers a column that is not hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Item', '', 'x', ''],
                        ['', '', '', ''],
                        ['', '', 'y', ''],
                    ],
                    mergeCells: { A1: [2, 1], C3: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.hideColumn(2);

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[2][2].element.style.display).to.equal('');
        expect(worksheet.records[1][1].element.style.display).to.equal('none');
        expect(worksheet.records[1][2].element.style.display).to.equal('none');

        // The visible first cell spans only the columns that are not hidden, so the next cells keep their columns
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('1');
        expect(worksheet.records[2][2].element.getAttribute('colspan')).to.equal('1');

        worksheet.showColumn([1, 2]);

        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('2');
        expect(worksheet.records[2][2].element.getAttribute('colspan')).to.equal('2');
    });

    it('hideColumn hides a merged range when each column that it covers is hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Item', '', 'x'],
                        ['', '', ''],
                    ],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        worksheet.hideColumn([0, 1]);

        expect(worksheet.records[0][0].element.style.display).to.equal('none');
    });

    it('setMerge gives the first cell only the columns that are not hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2] }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.setMerge('A1', 3, 1);

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('2');

        worksheet.showColumn(1);

        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('3');
    });

    it('keeps a merged range visible when one of its columns has the type hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['Item', 'b', 'c']],
                    columns: [{ type: 'hidden' }, {}, {}],
                    mergeCells: { A1: [2, 1] },
                },
            ],
        });

        const worksheet = instance[0];

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('1');
    });

    it('removeMerge keeps the cells of a hidden column hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: [['Item', '', 'x', 'y']], mergeCells: { A1: [2, 1] } }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(0);
        worksheet.removeMerge('A1');

        expect(worksheet.records[0][0].element.style.display).to.equal('none');
        expect(worksheet.records[0][1].element.style.display).to.equal('');

        worksheet.setMerge('A1', 2, 1);
        worksheet.hideColumn(1);
        worksheet.showColumn(0);
        worksheet.removeMerge('A1');

        expect(worksheet.records[0][0].element.style.display).to.equal('');
        expect(worksheet.records[0][1].element.style.display).to.equal('none');
    });

    it('selects each column of a merged range when one of its columns is hidden', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [4, 2], mergeCells: { A1: [3, 1] } }],
        });

        const worksheet = instance[0];

        worksheet.hideColumn(1);
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);

        expect(worksheet.headers.map((header) => header.classList.contains('selected'))).to.eql([true, false, true, false]);

        worksheet.resetSelection();

        expect(worksheet.headers.map((header) => header.classList.contains('selected'))).to.eql([false, false, false, false]);
    });
});
