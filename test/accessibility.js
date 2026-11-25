const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const createData = function (rows, columns) {
    const data = [];
    for (let j = 0; j < rows; j++) {
        const row = [];
        for (let i = 0; i < columns; i++) {
            row.push(j + '-' + i);
        }
        data.push(row);
    }
    return data;
};

// Returns the aria-rowindex of the row of each data cell and the aria-colindex of the cell, as 'row:column'
const getIndexes = function (worksheet) {
    return worksheet.records.map(function (row, y) {
        return row.map(function (record) {
            return worksheet.rows[y].element.getAttribute('aria-rowindex') + ':' + record.element.getAttribute('aria-colindex');
        });
    });
};

const getExpectedIndexes = function (rows, columns, headerRows) {
    const result = [];
    for (let j = 0; j < rows; j++) {
        const row = [];
        for (let i = 0; i < columns; i++) {
            row.push(j + 1 + (headerRows || 1) + ':' + (i + 2));
        }
        result.push(row);
    }
    return result;
};

describe('Grid roles for assistive technology', () => {
    it('Start the worksheet with the roles of the grid pattern', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 2] }],
        });
        const worksheet = instance[0];

        expect(worksheet.table.getAttribute('role')).to.equal('grid');
        expect(worksheet.thead.getAttribute('role')).to.equal('rowgroup');
        expect(worksheet.tbody.getAttribute('role')).to.equal('rowgroup');

        expect(worksheet.headerContainer.getAttribute('role')).to.equal('row');
        expect(worksheet.headerContainer.getAttribute('aria-rowindex')).to.equal('1');
        expect(worksheet.headerContainer.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.headers.map((td) => td.getAttribute('role'))).to.eql(['columnheader', 'columnheader', 'columnheader']);
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3', '4']);

        const row = worksheet.rows[1].element;
        expect(row.getAttribute('role')).to.equal('row');
        expect(row.getAttribute('aria-rowindex')).to.equal('3');
        expect(row.children[0].getAttribute('role')).to.equal('rowheader');
        expect(row.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.records[1][2].element.getAttribute('aria-colindex')).to.equal('4');

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('3');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');
    });

    it('Give the full size of the worksheet with lazy loading', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(300, 3), lazyLoading: true, tableOverflow: true, tableHeight: '300px' }],
        });
        const worksheet = instance[0];

        expect(worksheet.tbody.children.length).to.equal(100);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('301');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');

        // Go to the last page
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.last(false, true);

        const firstRow = worksheet.tbody.children[0];
        const lastRow = worksheet.tbody.children[worksheet.tbody.children.length - 1];
        expect(firstRow.getAttribute('data-y')).to.equal('200');
        expect(firstRow.getAttribute('aria-rowindex')).to.equal('202');
        expect(lastRow.getAttribute('aria-rowindex')).to.equal('301');
    });

    it('Give the full size of the worksheet with pagination', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(10, 2), pagination: 3 }],
        });
        const worksheet = instance[0];

        expect(worksheet.tbody.children.length).to.equal(3);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('11');

        worksheet.page(2);

        const rows = Array.prototype.map.call(worksheet.tbody.children, (tr) => tr.getAttribute('aria-rowindex'));
        expect(rows).to.eql(['8', '9', '10']);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('11');
    });

    it('Keep the row indexes after insertRow, deleteRow, moveRow and undo', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(4, 2) }],
        });
        const worksheet = instance[0];

        worksheet.insertRow(2, 0, true);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('7');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(6, 2));
        expect(worksheet.records[2][0].element.textContent).to.equal('0-0');

        worksheet.deleteRow(0, 3);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('4');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 2));

        worksheet.moveRow(0, 2);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 2));

        worksheet.undo();
        worksheet.undo();
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('7');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(6, 2));
    });

    it('Keep the column indexes after insertColumn, deleteColumn and moveColumn', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 3) }],
        });
        const worksheet = instance[0];

        worksheet.insertColumn(2, 0, true);
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('6');
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3', '4', '5', '6']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 5));

        worksheet.deleteColumn(0, 3);
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('3');
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2));

        worksheet.moveColumn(0, 1);
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2));
    });

    it('Keep the column indexes of merged cells after insertColumn', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 3), mergeCells: { B1: [2, 2] } }],
        });
        const worksheet = instance[0];

        worksheet.insertColumn(1, 0, true);

        expect(worksheet.getMerge('C1')).to.eql([2, 2]);
        expect(worksheet.records[0][2].element.getAttribute('aria-colindex')).to.equal('4');
        // The cells under the merged cell
        const cells = worksheet.options.mergeCells.C1[2];
        expect(cells.map((td) => td.getAttribute('data-x') + ':' + td.getAttribute('aria-colindex'))).to.eql(['3:5', '2:4', '3:5']);
    });

    it('Keep the row indexes after orderBy', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: [['c'], ['a'], ['b']] }],
        });
        const worksheet = instance[0];

        worksheet.orderBy(0, 0);

        expect(worksheet.getColumnData(0)).to.eql(['a', 'b', 'c']);
        expect(Array.prototype.map.call(worksheet.tbody.children, (tr) => tr.getAttribute('aria-rowindex'))).to.eql(['2', '3', '4']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 1));
    });

    it('Keep the indexes of hidden rows and columns', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(3, 3) }],
        });
        const worksheet = instance[0];

        worksheet.hideRow(1);
        worksheet.hideColumn(1);

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('4');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');
        expect(worksheet.rows[2].element.getAttribute('aria-rowindex')).to.equal('4');
        expect(worksheet.records[2][2].element.getAttribute('aria-colindex')).to.equal('4');
    });

    it('Count the rows of nested headers and footers', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: createData(2, 2),
                    nestedHeaders: [[{ title: 'Group', colspan: 2 }]],
                    footers: [['Total', '']],
                },
            ],
        });
        const worksheet = instance[0];
        const nested = worksheet.thead.children[0];
        const footer = worksheet.table.querySelector('tfoot > tr');

        expect(nested.getAttribute('role')).to.equal('row');
        expect(nested.getAttribute('aria-rowindex')).to.equal('1');
        expect(nested.children[1].getAttribute('role')).to.equal('columnheader');
        expect(worksheet.headerContainer.getAttribute('aria-rowindex')).to.equal('2');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2, 2));
        expect(footer.getAttribute('role')).to.equal('row');
        expect(footer.getAttribute('aria-rowindex')).to.equal('5');
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('5');

        worksheet.insertRow();

        expect(footer.getAttribute('aria-rowindex')).to.equal('6');
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('6');
    });

    it('Keep the column indexes of nested header, filter and footer cells after hideColumn', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: createData(2, 4),
                    filters: true,
                    nestedHeaders: [
                        [
                            { title: 'G1', colspan: 2 },
                            { title: 'G2', colspan: 2 },
                        ],
                    ],
                    footers: [['w', 'x', 'y', 'z']],
                },
            ],
        })[0];
        const nested = worksheet.thead.children[0];
        const footer = worksheet.table.querySelector('tfoot > tr');

        worksheet.hideColumn(1);

        // The cells of each row in the column D have the aria-colindex of the header of the column D
        const expected = worksheet.headers[3].getAttribute('aria-colindex');
        expect(expected).to.equal('5');
        expect(worksheet.filter.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(footer.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(worksheet.records[0][3].element.getAttribute('aria-colindex')).to.equal(expected);
        // The nested header G2 starts in the column C
        expect(nested.children[2].getAttribute('aria-colindex')).to.equal('4');
        expect(nested.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.filter.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(footer.children[0].getAttribute('aria-colindex')).to.equal('1');

        worksheet.deleteColumn(0);

        expect(footer.children[3].getAttribute('aria-colindex')).to.equal(worksheet.headers[2].getAttribute('aria-colindex'));
        expect(footer.children[3].getAttribute('aria-colindex')).to.equal('4');
    });

    it('Update the size of the grid after setData', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 2) }],
        });
        const worksheet = instance[0];

        worksheet.setData(createData(5, 2));

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('6');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(5, 2));
    });
});
