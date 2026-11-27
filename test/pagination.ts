import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

describe('Use pagination', () => {
    it('shows the page number and the number of pages', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 10], pagination: 3 }],
        })[0];

        worksheet.page(1);

        expect(worksheet.pagination!.children[0].textContent).to.equal('Showing page 2 of 4 entries');
    });

    it('Start the worksheet with pagination', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3, 4, 5],
                        [6, 7, 8, 9, 10],
                        [11, 12, 13, 14, 15],
                        [16, 17, 18, 19, 20],
                        [21, 22, 23, 24, 25],
                        [26, 27, 28, 29, 30],
                        [31, 32, 33, 34, 35],
                        [36, 37, 38, 39, 40],
                        [41, 42, 43, 44, 45],
                        [46, 47, 48, 49, 50],
                    ],
                    pagination: 3,
                },
            ],
        });

        expect(instance[0].quantiyOfPages()).to.equal(4);

        const bodyTag = root.querySelector('tbody')!;

        expect(bodyTag.children.length).to.equal(3);

        expect(bodyTag.children[0].getAttribute('data-y')).to.equal('0');
        expect(bodyTag.children[1].getAttribute('data-y')).to.equal('1');
        expect(bodyTag.children[2].getAttribute('data-y')).to.equal('2');
    });

    it('page method', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3, 4, 5],
                        [6, 7, 8, 9, 10],
                        [11, 12, 13, 14, 15],
                        [16, 17, 18, 19, 20],
                        [21, 22, 23, 24, 25],
                        [26, 27, 28, 29, 30],
                        [31, 32, 33, 34, 35],
                        [36, 37, 38, 39, 40],
                        [41, 42, 43, 44, 45],
                        [46, 47, 48, 49, 50],
                    ],
                    pagination: 3,
                },
            ],
        });

        instance[0].page(2);

        expect(instance[0].quantiyOfPages()).to.equal(4);

        const bodyTag = root.querySelector('tbody')!;

        expect(bodyTag.children.length).to.equal(3);

        expect(bodyTag.children[0].getAttribute('data-y')).to.equal('6');
        expect(bodyTag.children[1].getAttribute('data-y')).to.equal('7');
        expect(bodyTag.children[2].getAttribute('data-y')).to.equal('8');
    });
});

describe('Row changes show the rows of the page again (#223, #224)', () => {
    const confirm = globalThis.confirm;

    beforeEach(() => {
        globalThis.confirm = () => true;
    });

    afterEach(() => {
        globalThis.confirm = confirm;
    });

    function create(pageNumber: number) {
        const data = Array.from({ length: 9 }, (row, y) => [`A${y + 1}`, `B${y + 1}`]);
        const worksheet = jspreadsheet(root, { worksheets: [{ data, pagination: 3 }] })[0];
        worksheet.page(pageNumber);
        return worksheet;
    }

    function shown(worksheet: ReturnType<typeof create>) {
        return Array.from(worksheet.tbody.children).map((tr) => tr.children[1].textContent);
    }

    it('shows the last page when a delete removes the current page', () => {
        const worksheet = create(2);

        worksheet.deleteRow(0, 3);

        expect(worksheet.view.pageNumber).to.equal(1);
        expect(shown(worksheet)).to.eql(['A7', 'A8', 'A9']);
    });

    it('shows the last page when the undo of an insert removes the current page', () => {
        const worksheet = create(2);
        worksheet.insertRow(3, 8);
        worksheet.page(3);

        worksheet.undo();

        expect(worksheet.view.pageNumber).to.equal(2);
        expect(shown(worksheet)).to.eql(['A7', 'A8', 'A9']);
    });

    it('shows the next rows after a delete of rows on an earlier page (#223)', () => {
        const worksheet = create(1);

        worksheet.deleteRow(0, 2);

        expect(Array.from(worksheet.tbody.children).map((tr) => tr.getAttribute('data-y'))).to.eql(['3', '4', '5']);
        expect(shown(worksheet)).to.eql(['A6', 'A7', 'A8']);
    });

    it('redoes a delete of rows that are not on the page (#224)', () => {
        const worksheet = create(0);
        worksheet.deleteRow(0, 2);
        worksheet.undo();
        worksheet.page(2);

        worksheet.redo();

        expect(worksheet.options.data.length).to.equal(7);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('A3');
        expect(shown(worksheet)).to.eql(['A9']);
    });

    it('undoes and redoes an insert of rows that are not on the page', () => {
        const worksheet = create(2);

        worksheet.insertRow(2, 0);
        worksheet.undo();
        expect(worksheet.options.data.length).to.equal(9);
        expect(shown(worksheet)).to.eql(['A7', 'A8', 'A9']);

        worksheet.redo();
        expect(worksheet.options.data.length).to.equal(11);
        expect(shown(worksheet)).to.eql(['A5', 'A6', 'A7']);
    });

    it('moves a row between two other pages, and shows the rows of the page', () => {
        const worksheet = create(1);

        worksheet.moveRow(7, 0);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('A8');
        expect(shown(worksheet)).to.eql(['A3', 'A4', 'A5']);

        worksheet.moveRow(0, 8);
        expect(shown(worksheet)).to.eql(['A4', 'A5', 'A6']);

        worksheet.undo();
        worksheet.undo();
        expect(worksheet.getColumnData(0)).to.eql(['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9']);
        expect(shown(worksheet)).to.eql(['A4', 'A5', 'A6']);
    });

    it('moves a row with a filter, and the filter shows the row at its new place', () => {
        const data = Array.from({ length: 10 }, (row, y) => [`${y}`, `b${y}`]);
        const worksheet = jspreadsheet(root, { worksheets: [{ data, filters: true, pagination: 3 }] })[0];
        worksheet.view.filters[0] = ['0', '2', '4', '6', '8'];
        worksheet.view.closeFilter(0);

        worksheet.moveRow(0, 5);
        expect(worksheet.view.results).to.eql([1, 3, 5, 6, 8]);
        expect(Array.from(worksheet.tbody.children).map((tr) => tr.children[1].textContent)).to.eql(['2', '4', '0']);

        worksheet.moveRow(1, 6);
        expect(worksheet.view.results).to.eql([2, 4, 5, 6, 8]);
        expect(Array.from(worksheet.tbody.children).map((tr) => tr.children[1].textContent)).to.eql(['4', '0', '6']);
    });
});

describe('Insert and delete change the results of a filter (#234)', () => {
    function create(options: Record<string, unknown> = {}) {
        const data = Array.from({ length: 10 }, (row, y) => [`${y}`, `b${y}`]);
        const worksheet = jspreadsheet(root, { worksheets: [{ data, filters: true, ...options }] })[0];
        worksheet.view.filters[0] = ['0', '2', '4', '6', '8'];
        worksheet.view.closeFilter(0);
        return worksheet;
    }

    function shown(worksheet: ReturnType<typeof create>) {
        return Array.from(worksheet.tbody.children).map((tr) => tr.children[1].textContent);
    }

    it('keeps the rows of the filter after a delete, and the page shows them', () => {
        const worksheet = create({ pagination: 3 });

        worksheet.deleteRow(0, 1);
        expect(worksheet.view.results).to.eql([1, 3, 5, 7]);
        expect(shown(worksheet)).to.eql(['2', '4', '6']);
    });

    it('shows a new row at its place after an insert, and keeps the rows of the filter', () => {
        const worksheet = create({ pagination: 3 });

        worksheet.insertRow(1, 0, true);
        expect(worksheet.view.results).to.eql([0, 1, 3, 5, 7, 9]);
        expect(shown(worksheet)).to.eql(['', '0', '2']);
    });

    it('puts new rows in the body between the rows of the filter, and undo and redo change the results', () => {
        const worksheet = create();

        worksheet.insertRow(2, 3);
        expect(worksheet.view.results).to.eql([0, 2, 4, 5, 6, 8, 10]);
        expect(shown(worksheet)).to.eql(['0', '2', '', '', '4', '6', '8']);

        worksheet.undo();
        expect(worksheet.view.results).to.eql([0, 2, 4, 6, 8]);
        expect(shown(worksheet)).to.eql(['0', '2', '4', '6', '8']);

        worksheet.redo();
        expect(worksheet.view.results).to.eql([0, 2, 4, 5, 6, 8, 10]);
        expect(shown(worksheet)).to.eql(['0', '2', '', '', '4', '6', '8']);
    });

    it('removes the deleted rows from the results, and shows the rows that an undo gives back', () => {
        const worksheet = create();

        worksheet.deleteRow(2, 2);
        expect(worksheet.view.results).to.eql([0, 2, 4, 6]);
        expect(shown(worksheet)).to.eql(['0', '4', '6', '8']);

        worksheet.undo();
        expect(worksheet.view.results).to.eql([0, 2, 3, 4, 6, 8]);
        expect(shown(worksheet)).to.eql(['0', '2', '3', '4', '6', '8']);

        worksheet.redo();
        expect(worksheet.view.results).to.eql([0, 2, 4, 6]);
        expect(shown(worksheet)).to.eql(['0', '4', '6', '8']);
    });

    it('keeps the filter when a delete removes all its rows, and shows a new row', () => {
        const worksheet = create();
        worksheet.view.filters[0] = ['4'];
        worksheet.view.closeFilter(0);

        worksheet.deleteRow(4, 1);
        expect(worksheet.view.results).to.eql([]);
        expect(shown(worksheet)).to.eql([]);

        worksheet.insertRow(1, 0);
        expect(worksheet.view.results).to.eql([1]);
        expect(shown(worksheet)).to.eql(['']);
    });

    it('shows a row that the filter hid before the delete after the undo', () => {
        const worksheet = create();

        worksheet.deleteRow(3, 1);
        worksheet.undo();
        expect(worksheet.rows[3].element.style.display).to.equal('');
        expect(shown(worksheet)).to.eql(['0', '2', '3', '4', '6', '8']);
    });
});
