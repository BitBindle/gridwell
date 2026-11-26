import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

function rows(count: number) {
    const data: string[][] = [];
    for (let j = 0; j < count; j++) {
        data.push([`${j}`]);
    }
    return data;
}

describe('Lazy loading', () => {
    it('loads the page of the selected row when the keyboard moves from a row that is not attached', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // Row 251 is not attached
        worksheet.updateSelectionFromCoords(0, 250, 0, 250);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 251, 0, 251]);
        expect(worksheet.tbody.contains(worksheet.records[251][0].element)).to.equal(true);

        worksheet.updateSelectionFromCoords(0, 250, 0, 250);
        worksheet.up();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 249, 0, 249]);
        expect(worksheet.tbody.contains(worksheet.records[249][0].element)).to.equal(true);
    });

    it('moves up to the first attached row without an error', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // A move near the last attached row attaches the next rows and removes the first rows
        worksheet.updateSelectionFromCoords(0, 80, 0, 80);
        worksheet.down();

        const first = parseInt((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));
        expect(first).to.be.above(0);

        worksheet.updateSelectionFromCoords(0, first + 1, 0, first + 1);
        worksheet.up();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, first, 0, first]);
    });

    it('loads the page of the selected result when the keyboard moves with a search', () => {
        const data: string[][] = [];
        for (let j = 0; j < 600; j++) {
            data.push([(j % 2 ? 'b' : 'a') + j]);
        }

        const instance = jspreadsheet(root, {
            worksheets: [{ data, search: true, lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // The search keeps the even rows. Rows 302 and 452 are not attached.
        worksheet.search('a');

        worksheet.updateSelectionFromCoords(0, 300, 0, 300);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 302, 0, 302]);
        expect(worksheet.tbody.contains(worksheet.records[302][0].element)).to.equal(true);

        worksheet.updateSelectionFromCoords(0, 450, 0, 450);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 452, 0, 452]);
        expect(worksheet.tbody.contains(worksheet.records[452][0].element)).to.equal(true);
    });
});
