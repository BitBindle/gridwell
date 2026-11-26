import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

function grid(rows: number, columns: number, formula?: boolean) {
    const data: (string | number)[][] = [];
    for (let j = 0; j < rows; j++) {
        const row: (string | number)[] = [];
        for (let i = 0; i < columns; i++) {
            row.push(`${i}:${j}`);
        }
        data.push(row);
    }
    if (formula) {
        data[0][0] = '=C5';
    }
    return data;
}

// Each row, each cell and each header has the position of its index, also after a change in the middle
function positionsMatch(worksheet: Worksheet) {
    for (let j = 0; j < worksheet.rows.length; j++) {
        const row = worksheet.rows[j].element;
        if (row.getAttribute('data-y') !== `${j}` || row.children[0].textContent !== `${j + 1}`) {
            return `row ${j}`;
        }
        for (let i = 0; i < worksheet.headers.length; i++) {
            const cell = worksheet.records[j][i].element;
            if (cell.getAttribute('data-x') !== `${i}` || cell.getAttribute('data-y') !== `${j}`) {
                return `cell ${i},${j}`;
            }
        }
    }
    for (let i = 0; i < worksheet.headers.length; i++) {
        if (worksheet.headers[i].getAttribute('data-x') !== `${i}`) {
            return `header ${i}`;
        }
    }
    return 'match';
}

describe('Positions after an insert, a move or a delete', () => {
    const changes: [string, (w: Worksheet) => void][] = [
        ['insertRow in the middle', (w) => w.insertRow(2, 6)],
        ['insertRow at the top', (w) => w.insertRow(1, 0, true)],
        ['deleteRow in the middle', (w) => w.deleteRow(6, 2)],
        ['moveRow', (w) => w.moveRow(8, 3)],
        ['insertColumn in the middle', (w) => w.insertColumn(2, 3)],
        ['deleteColumn in the middle', (w) => w.deleteColumn(3, 2)],
        ['moveColumn', (w) => w.moveColumn(5, 1)],
    ];

    for (const [name, change] of changes) {
        it(`gives each row, cell and header its position after ${name}`, () => {
            for (const formula of [false, true]) {
                const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid(10, 6, formula) }] });

                change(worksheet);

                expect(positionsMatch(worksheet), `${name}, formula ${formula}`).to.equal('match');
            }
        });
    }

    it('keeps the references of a formula and the meta information on their cells', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid(10, 6, true), meta: { C5: { note: 'x' } } }] });

        worksheet.insertRow(1, 1);
        expect(worksheet.getValue('A1')).to.equal('=C6');
        expect(worksheet.getMeta('C6')).to.eql({ note: 'x' });

        worksheet.insertColumn(1, 1);
        expect(worksheet.getValue('A1')).to.equal('=D6');
        expect(worksheet.getMeta('D6')).to.eql({ note: 'x' });

        worksheet.moveRow(5, 0);
        expect(worksheet.getValue('A2')).to.equal('=D1');
    });

    it('moves a merged range after an insert in a worksheet without formulas', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid(10, 6), mergeCells: { B6: [2, 2] } }] });

        worksheet.insertRow(1, 2);

        expect(worksheet.getMerge()).to.eql({ B7: [2, 2] });
        expect(worksheet.records[6][1].element.getAttribute('rowspan')).to.equal('2');
    });

    it('keeps a merged range that ends on the row or column of an insert after it', () => {
        const cases: [Record<string, [number, number]>, (w: Worksheet) => void, Record<string, [number, number]>][] = [
            [{ A2: [1, 2] }, (w) => w.insertRow(1, 2), { A2: [1, 2] }],
            [{ B3: [2, 3] }, (w) => w.insertRow(2, 4), { B3: [2, 3] }],
            [{ A2: [2, 1] }, (w) => w.insertColumn(1, 1), { A2: [2, 1] }],
            [{ B2: [2, 2] }, (w) => w.insertColumn(1, 2), { B2: [2, 2] }],
            // A range before the start leads to a second range before it
            [{ C2: [1, 2], A3: [2, 2] }, (w) => w.insertRow(1, 3), { C2: [1, 2], A3: [2, 2] }],
            [{ B3: [2, 1], C1: [2, 2] }, (w) => w.insertColumn(1, 3), { B3: [2, 1], C1: [2, 2] }],
        ];
        for (const [mergeCells, change, expected] of cases) {
            const data = grid(10, 6);
            data[0][5] = '=A2';
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data, mergeCells }] });

            change(worksheet);

            const names = Object.keys(expected);
            expect(worksheet.getMerge(), JSON.stringify(mergeCells)).to.deep.equal(expected);
            // An inserted column moves the formula, but its reference stays
            expect(worksheet.getRowData(0), JSON.stringify(mergeCells)).to.include('=A2');

            // Each range splits as one range
            for (const name of names) {
                worksheet.removeMerge(name);
            }
            expect(worksheet.getMerge()).to.eql({});
        }
    });

    it('moves the formula chain also when no cell holds a formula any more', () => {
        const changes: string[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: grid(10, 6) }],
            onchange: (instance, cell, x, y, value) => {
                changes.push(`${x},${y}=${value}`);
            },
        });
        worksheet.setValue('C3', '=A1+1');
        worksheet.setValue('C3', 5);

        worksheet.insertRow(1, 0, true);
        worksheet.setValue('C3', '=B1');
        changes.length = 0;
        worksheet.setValue('A1', 2);

        // The chain moved with the old cells, so C3 does not depend on A1 and is not calculated again
        expect(changes).to.eql(['0,0=2']);
    });
});
