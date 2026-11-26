import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

function rows(count: number) {
    return Array.from({ length: count }, (row, y) => [`${y}`, `b${y}`]);
}

function attached(worksheet: Worksheet) {
    return Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
}

function range(from: number, to: number) {
    return Array.from({ length: to - from + 1 }, (item, k) => from + k);
}

describe('loadRow', () => {
    it('fills the table with the rows from 10 rows above the row, and keeps the scroll position', () => {
        const events: [number, number][][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }],
            onrowsattached: (instance, ranges) => {
                events.push(ranges);
            },
        });
        worksheet.content.scrollTop = 7;

        worksheet.loadRow(500);

        expect(attached(worksheet)).to.eql(range(490, 589));
        expect(events).to.eql([[[490, 589]]]);
        expect(worksheet.content.scrollTop).to.equal(7);
    });

    it('keeps the frozen rows first, and gives only the new rows to onrowsattached', () => {
        const events: [number, number][][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true, freezeRows: 3 }],
            onrowsattached: (instance, ranges) => {
                events.push(ranges);
            },
        });

        worksheet.loadRow(500);
        expect(attached(worksheet)).to.eql([0, 1, 2, ...range(490, 586)]);

        worksheet.loadRow(520);
        expect(attached(worksheet)).to.eql([0, 1, 2, ...range(510, 606)]);
        expect(events).to.eql([[[490, 586]], [[587, 606]]]);
    });

    it('fills the table with the last rows for a row near the end', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }] });

        worksheet.loadRow(995);

        expect(attached(worksheet)).to.eql(range(900, 999));
    });

    it('uses the rows of the search results, from the first result at or after the row', () => {
        const data = rows(1000).map(([a, b], y) => [y % 2 ? 'odd' : a, b]);
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data, lazyLoading: true, tableOverflow: true, search: true }] });
        worksheet.search('odd');

        worksheet.loadRow(600);

        // Row 601 is the result at index 300, so the table starts at the result at index 290
        expect(attached(worksheet)).to.eql(range(290, 389).map((k) => 2 * k + 1));
    });

    it('saves the open editor of a row that leaves the table', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }] });
        const cell = worksheet.getCell('A2');
        worksheet.openEditor(cell);
        (cell.children[0] as HTMLInputElement).value = 'new';

        worksheet.loadRow(500);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValue('A2')).to.equal('new');
    });

    it('places the selection corner again, and does nothing for a row index that is not a number', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }] });
        worksheet.updateSelectionFromCoords(0, 2, 0, 2);
        worksheet.corner.style.top = '1234px';

        worksheet.loadRow(500);
        expect(worksheet.corner.style.top).to.not.equal('1234px');

        for (const value of [null, '200', NaN, Infinity]) {
            worksheet.loadRow(value as unknown as number);
            expect(attached(worksheet)[0]).to.equal(490);
        }
    });

    it('does nothing without lazy loading', () => {
        let events = 0;
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(300) }],
            onrowsattached: () => {
                events++;
            },
        });

        worksheet.loadRow(200);

        expect(attached(worksheet)).to.eql(range(0, 299));
        expect(events).to.equal(0);
    });
});

describe('onrowsattached', () => {
    function build(options: Partial<Jspreadsheet.WorksheetOptions> = {}) {
        const events: [number, number][][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true, ...options }],
            onrowsattached: (instance, ranges) => {
                events.push(ranges);
            },
        });
        return { worksheet, events };
    }

    it('gives the rows that a scroll down attaches', async () => {
        // The grid loads rows 100 ms after a scroll, one load at a time. Let the loads of the tests before end.
        await new Promise((resolve) => setTimeout(resolve, 150));
        const { worksheet, events } = build();

        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
        await new Promise((resolve) => setTimeout(resolve, 150));

        expect(events).to.eql([[[100, 129]]]);
        expect(attached(worksheet)).to.eql(range(30, 129));
    });

    it('gives the rows of the results of a search and of its reset', () => {
        const data = rows(1000).map(([a, b], y) => [y % 2 ? 'odd' : a, b]);
        const { worksheet, events } = build({ data, search: true });

        worksheet.search('odd');
        expect(events).to.eql([
            [
                [101, 101],
                [103, 103],
            ].concat(range(52, 99).map((k) => [2 * k + 1, 2 * k + 1])),
        ]);

        events.length = 0;
        worksheet.resetSearch();
        expect(events).to.eql([
            [
                [0, 0],
                [2, 2],
            ].concat(range(2, 49).map((k) => [2 * k, 2 * k])),
        ]);
    });

    it('gives the rows that a sort and its undo attach, also while the events are ignored', () => {
        const { worksheet, events } = build();

        worksheet.orderBy(0, 1);
        expect(events.length).to.equal(1);

        events.length = 0;
        worksheet.undo();
        // The descending sort of the texts showed the rows "99" to "91" also, so they were in the table
        expect(attached(worksheet)).to.eql(range(0, 99));
        expect(events).to.eql([[[0, 90]]]);
    });

    it('gives no rows that were in the table before the load', () => {
        const { worksheet, events } = build();
        worksheet.loadRow(150);
        events.length = 0;

        // The selection of a row outside the table loads its page
        worksheet.updateSelectionFromCoords(0, 120, 0, 120);
        worksheet.loadRow(120);

        expect(events).to.eql([[[110, 139]]]);
    });
});
