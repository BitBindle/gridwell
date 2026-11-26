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
});
