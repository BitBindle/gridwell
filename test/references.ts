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

    it('shows the result of each formula text after each change, as a new worksheet with the same data does', () => {
        const numbers = () => Array.from({ length: 8 }, (row, j) => [`${j + 1}`, `${(j + 1) * 10}`, '', '']);
        const data = numbers();
        data[0][2] = '=A5+B6';
        data[1][3] = '=SUM(A1:A8)';
        data[7][2] = '=A8*2';
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data }], updateReferences: false });
        const steps: [string, () => void][] = [
            ['insertRow', () => worksheet.insertRow(1, 2)],
            ['deleteRow', () => worksheet.deleteRow(0, 1)],
            ['deleteRow at the end', () => worksheet.deleteRow(worksheet.rows.length - 1, 1)],
            ['moveRow', () => worksheet.moveRow(4, 0)],
            ['orderBy', () => worksheet.orderBy(1, 1)],
            ['undo', () => worksheet.undo()],
            ['insertColumn', () => worksheet.insertColumn(1, 0, true)],
            ['moveColumn', () => worksheet.moveColumn(0, 2)],
            ['deleteColumn', () => worksheet.deleteColumn(0, 1)],
        ];

        for (const [name, step] of steps) {
            step();
            const element = document.createElement('div');
            root.appendChild(element);
            const [fresh] = jspreadsheet(element, { worksheets: [{ data: worksheet.getData() }] });
            const shown = worksheet.getData(false, true);
            expect(shown, name).to.eql(fresh.getData(false, true));
        }
    });

    it('shows #ERROR for a formula that a move puts on a cell that its text names', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=A2'],
                        ['5', ''],
                    ],
                },
            ],
            updateReferences: false,
        });

        worksheet.moveColumn(1, 0);
        worksheet.moveRow(0, 1);

        expect(worksheet.getValue('A2')).to.equal('=A2');
        expect(worksheet.getValue('A2', true)).to.equal('#ERROR');
    });

    it('dispatches no change event and writes no formula text when it shows the new results', () => {
        const events: string[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['2', '=A2*2'],
                    ],
                },
            ],
            updateReferences: false,
            onbeforechange: () => {
                events.push('beforechange');
                return false;
            },
            onchange: () => {
                events.push('change');
            },
            oninsertrow: () => {
                events.push('insertrow');
            },
        });

        worksheet.insertRow(1, 0, true);

        expect(events).to.eql(['insertrow']);
        expect(worksheet.getValue('B3')).to.equal('=A2*2');
        expect(worksheet.getValue('B3', true)).to.equal('2');
    });
});
