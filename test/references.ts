import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

function grid() {
    const data: (string | number)[][] = [];
    for (let j = 0; j < 8; j++) {
        data.push([`${j}`, j, `c${j}`, `d${j}`]);
    }
    // C1 refers to A5 and to B6
    data[0][2] = '=A5+B6';
    return data;
}

describe('The updateReferences option', () => {
    const changes: [string, (w: Worksheet) => void, string][] = [
        ['insertRow', (w) => w.insertRow(1, 2), 'C1'],
        ['deleteRow', (w) => w.deleteRow(2, 1), 'C1'],
        ['moveRow', (w) => w.moveRow(4, 0), 'C2'],
        ['insertColumn', (w) => w.insertColumn(1, 0, true), 'D1'],
        ['deleteColumn', (w) => w.deleteColumn(0, 1), 'B1'],
        ['moveColumn', (w) => w.moveColumn(0, 3), 'B1'],
    ];

    for (const [name, change, cell] of changes) {
        it(`keeps the formula text after ${name} when updateReferences is false`, () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid() }], updateReferences: false });

            change(worksheet);

            expect(worksheet.getValue(cell)).to.equal('=A5+B6');
        });
    }

    it('keeps the formula text after a sort when updateReferences is false', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid() }], updateReferences: false });

        worksheet.orderBy(1, 1);

        expect(worksheet.getData().flat()).to.include('=A5+B6');
    });

    it('changes the references by default', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid() }] });

        worksheet.insertRow(1, 2);

        expect(worksheet.getValue('C1')).to.equal('=A6+B7');
    });

    it('calculates the moved formula again when a cell that its text names changes', () => {
        const changed: string[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: grid() }],
            updateReferences: false,
            formulaEngine: (instance, expression) => `${expression} calculated`,
            onchange: (instance, cell, x, y) => {
                changed.push(`${x},${y}`);
            },
        });

        // The old A5 moves to A6, and the formula text still names A5
        worksheet.insertRow(1, 2);
        changed.length = 0;
        worksheet.setValue('A5', 'new');
        expect(changed).to.include('2,0');

        // The formula moves from C1 to C3, and its text still names A5 and B6
        worksheet.moveRow(0, 2);
        changed.length = 0;
        worksheet.setValue('A5', 'newer');
        expect(changed).to.include('2,2');
    });

    it('keeps the meta information on its cell when updateReferences is false', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: grid(), meta: { D5: { note: 'x' } } }], updateReferences: false });

        worksheet.insertRow(1, 2);

        expect(worksheet.getMeta('D6')).to.eql({ note: 'x' });
    });
});
