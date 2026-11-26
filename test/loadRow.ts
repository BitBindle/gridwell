import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

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
    it('gives the rows that a scroll down and a scroll up attach', () => {
        const events: [number, number][][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }],
            onrowsattached: (instance, ranges) => {
                events.push(ranges);
            },
        });

        worksheet.view.loadDown();
        expect(events).to.eql([[[100, 129]]]);
        expect(attached(worksheet)).to.eql(range(30, 129));

        worksheet.view.loadUp();
        expect(events[1]).to.eql([[0, 29]]);
    });

    it('gives the rows of a page that loads for the selection', () => {
        const events: [number, number][][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(1000), lazyLoading: true, tableOverflow: true }],
            onrowsattached: (instance, ranges) => {
                events.push(ranges);
            },
        });

        worksheet.view.loadPage(4);

        expect(events).to.eql([[[400, 499]]]);
    });
});
