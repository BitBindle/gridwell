import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

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

    it('shows a new result as a change of the cell shows it: secured formulas, the render of the column, a custom type', () => {
        const custom = {
            createCell: (cell: HTMLTableCellElement, value: unknown) => {
                cell.innerHTML = `<i>${value}</i>`;
                return cell;
            },
            setValue: (cell: HTMLTableCellElement, value: unknown) => {
                cell.innerHTML = `<i>${value}</i>`;
            },
            openEditor: () => undefined,
            closeEditor: () => undefined,
        };
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=a2*2', '=A2', '=C1', '=A2*3'],
                        ['2', '', '', '', ''],
                        ['3', '', '', '', ''],
                    ],
                    secureFormulas: true,
                    columns: [{}, {}, { type: custom as unknown as Jspreadsheet.CustomEditor }, {}, { render: (td: HTMLTableCellElement) => td.prepend('R') }],
                },
            ],
            updateReferences: false,
        });
        const shown = () => [worksheet.getCell('B1').textContent, worksheet.getCell('C1').innerHTML, worksheet.getCell('E1').textContent];
        expect(shown()).to.eql(['4', '<i>=A2</i>', 'R6']);

        worksheet.insertRow(1, 0, true);

        // The old row 1 is row 2 now, so A2 is 1
        expect(worksheet.getValueFromCoords(1, 1)).to.equal('=a2*2');
        expect([worksheet.getCell('B2').textContent, worksheet.getCell('C2').innerHTML, worksheet.getCell('E2').textContent]).to.eql(['2', '<i>=A2</i>', 'R3']);
    });
});

describe('The formulas and the chain after an insert, a delete, and their undo (#197, #199)', () => {
    it('keeps the references of a formula in a row that the undo of a delete brings back', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '10', '=A3+B3'],
                        ['2', '20', ''],
                        ['3', '30', ''],
                        ['4', '40', ''],
                    ],
                },
            ],
        });

        worksheet.deleteRow(0, 1);
        worksheet.undo();

        expect(worksheet.getValue('C1')).to.equal('=A3+B3');
        expect(worksheet.formulas.chain).to.eql({ A3: ['C1'], B3: ['C1'] });
        worksheet.setValue('A3', '100');
        expect(worksheet.getValue('C1', true)).to.equal('130');
    });

    it('keeps the references of a formula in a column that the undo of a delete brings back', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['=B2+C2', '2', '3'],
                        ['', '4', '5'],
                    ],
                },
            ],
        });

        worksheet.deleteColumn(0, 1);
        worksheet.undo();

        expect(worksheet.getValue('A1')).to.equal('=B2+C2');
        worksheet.setValue('B2', '10');
        expect(worksheet.getValue('A1', true)).to.equal('15');
    });

    it('moves the references of the formula of a new row once, and shows the value and keeps the chain of the moved text', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['2', '=A1'],
                        ['3', ''],
                    ],
                },
            ],
        });

        worksheet.insertRow(['9', '=A3*2'], 0, true);

        // As in version 5, the formula names the cells at their positions before the insert
        expect(worksheet.getData()).to.eql([
            ['9', '=A4*2'],
            ['1', ''],
            ['2', '=A2'],
            ['3', ''],
        ]);
        expect(worksheet.getValue('B1', true)).to.equal('6');
        expect(worksheet.formulas.chain).to.eql({ A2: ['B3'], A4: ['B1'] });
        worksheet.undo();
        worksheet.redo();
        expect(worksheet.getValue('B1')).to.equal('=A4*2');
        worksheet.setValue('A4', '7');
        expect(worksheet.getValue('B1', true)).to.equal('14');
        expect(worksheet.getValue('B3', true)).to.equal('1');
    });

    it('does not throw on a change of a cell that a deleted formula named', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['2', ''],
                        ['3', '=A1*2'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(2, 1);

        expect(worksheet.formulas.chain).to.eql({});
        worksheet.setValue('A1', '5');
        expect(worksheet.getValue('A1')).to.equal('5');
    });

    it('does not calculate the cell that takes the name of a deleted formula cell (#239)', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['2', '=A1*2'],
                        ['3', '=A2*3'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(1, 1);

        // B2 is the old B3, whose formula named the deleted A2
        expect(worksheet.getValue('B2')).to.equal('=#REF!*3');
        expect(worksheet.getValue('B2', true)).to.equal('#REF!');
        expect(worksheet.formulas.chain).to.eql({});
    });

    it('gives #REF! to the references to a deleted cell, and keeps the dependents of the cell that takes its name (#239)', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=A3'],
                        ['2', '=A4'],
                        ['3', ''],
                        ['4', ''],
                    ],
                },
            ],
        });

        worksheet.deleteRow(2, 1);

        // B1 named the deleted A3. B2 names the old A4, which is A3 now.
        expect(worksheet.getData().map((row) => row[1])).to.eql(['=#REF!', '=A3', '']);
        worksheet.setValue('A3', '8');
        expect([worksheet.getValue('B1', true), worksheet.getValue('B2', true)]).to.eql(['#REF!', '8']);
    });

    it('removes the formula cells from the chain when a delete removes all rows', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    allowDeletingAllRows: true,
                    data: [
                        ['1', ''],
                        ['2', '=A1'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(0, 2);

        expect(worksheet.formulas.chain).to.eql({});
    });

    it('moves the references of the formula of a new column once, also when the redo adds the column again', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '2'],
                        ['3', '4'],
                    ],
                },
            ],
        });

        worksheet.insertColumn(['=B2*2', ''], 0, true);
        expect(worksheet.getValue('A1')).to.equal('=C2*2');
        expect(worksheet.getValue('A1', true)).to.equal('8');
        worksheet.undo();
        worksheet.redo();

        expect(worksheet.getValue('A1')).to.equal('=C2*2');
        worksheet.setValue('C2', '5');
        expect(worksheet.getValue('A1', true)).to.equal('10');
    });
});

describe('References to deleted cells (#239)', () => {
    const column = (worksheet: Worksheet, x: number, shown = false) =>
        Array.from({ length: worksheet.options.data.length }, (_, y) => worksheet.getValueFromCoords(x, y, shown));

    it('gives #REF! to a reference to a deleted row, and the undo gives back the formula texts', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['5', '=A1'],
                        ['7', '=A2'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(0, 1);
        expect(column(worksheet, 1)).to.eql(['=#REF!', '=A1']);
        expect(column(worksheet, 1, true)).to.eql(['#REF!', '5']);

        worksheet.undo();
        expect(column(worksheet, 1)).to.eql(['', '=A1', '=A2']);
        expect(column(worksheet, 1, true)).to.eql(['', '1', '5']);

        worksheet.redo();
        expect(column(worksheet, 1)).to.eql(['=#REF!', '=A1']);
        expect(column(worksheet, 1, true)).to.eql(['#REF!', '5']);

        // The formulas that the undo gives back are in the chain
        worksheet.undo();
        worksheet.setValue('A1', '3');
        expect(column(worksheet, 1, true)).to.eql(['', '3', '5']);
    });

    it('gives #REF! to a formula that names a cell with the result #REF!', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['2', '=A1*2'],
                        ['3', '=B2+1'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(0, 1);
        expect(column(worksheet, 1)).to.eql(['=#REF!*2', '=B1+1']);
        expect(column(worksheet, 1, true)).to.eql(['#REF!', '#REF!']);

        worksheet.undo();
        expect(column(worksheet, 1, true)).to.eql(['', '2', '3']);
    });

    it('ends a range at the cells that stay, and gives #REF! to a range whose cells all go', () => {
        const data = () => [
            ['1', '=SUM(A2:A4)'],
            ['2', '=A4:A2'],
            ['3', ''],
            ['4', ''],
            ['5', ''],
        ];
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: data() }] });

        worksheet.deleteRow(2, 1);
        expect(column(worksheet, 1)).to.eql(['=SUM(A2:A3)', '=A3:A2', '', '']);
        worksheet.undo();

        worksheet.deleteRow(3, 2);
        expect(column(worksheet, 1)).to.eql(['=SUM(A2:A3)', '=A3:A2', '']);
        worksheet.undo();

        worksheet.deleteRow(1, 1);
        expect(column(worksheet, 1)).to.eql(['=SUM(A2:A3)', '', '', '']);
        worksheet.undo();

        worksheet.deleteRow(2, 3);
        expect(column(worksheet, 1)).to.eql(['=SUM(A2:A2)', '=A2:A2']);
        worksheet.undo();
        expect(worksheet.getData()).to.eql(data());

        worksheet.deleteRow(1, 4);
        expect(column(worksheet, 1)).to.eql(['=SUM(#REF!)']);
        expect(worksheet.getValue('B1', true)).to.equal('#REF!');
        worksheet.undo();
        expect(worksheet.getData()).to.eql(data());
    });

    it('gives #REF! to a reference to a deleted column, and keeps the string literals', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: [['1', '2', '=A1+B1', '="A1"&B1']] }] });

        worksheet.deleteColumn(0, 1);
        expect(worksheet.getData()).to.eql([['2', '=#REF!+A1', '="A1"&A1']]);
        expect(worksheet.getValue('B1', true)).to.equal('#REF!');

        worksheet.undo();
        expect(worksheet.getData()).to.eql([['1', '2', '=A1+B1', '="A1"&B1']]);
        expect(worksheet.getValue('C1', true)).to.equal('3');

        worksheet.redo();
        expect(worksheet.getData()).to.eql([['2', '=#REF!+A1', '="A1"&A1']]);
    });

    it('gives #REF! to references with $ signs, keeps the $ signs of a range end, and keeps function names and other worksheets', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=$A3+A$3*10'],
                        ['2', '=SUM($A$1:A3)'],
                        ['3', ''],
                        ['4', '=$A$3'],
                        ['5', '=ATAN2(A5,A6)+LOG10(A4)'],
                        ['6', '=$A5&Sheet2!A3'],
                    ],
                },
            ],
        });
        const before = column(worksheet, 1);

        worksheet.deleteRow(2, 1);
        expect(column(worksheet, 1)).to.eql(['=#REF!+#REF!*10', '=SUM($A$1:A2)', '=#REF!', '=ATAN2(A4,A5)+LOG10(A3)', '=$A4&Sheet2!A3']);
        expect(worksheet.getValue('B1', true)).to.equal('#REF!');
        expect(worksheet.getValue('B3', true)).to.equal('#REF!');

        worksheet.undo();
        expect(column(worksheet, 1)).to.eql(before);
    });

    it('gives back on undo the formulas that the delete keeps but the undo renames, and ends ranges after row 9', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1', '2', '3', '=A1&"C1"', '=A1+Sheet2!C1', '=SUM(C$1:A1)']],
                },
            ],
        });
        const before = worksheet.getData()[0].slice(3);
        worksheet.deleteColumn(1, 1);
        worksheet.undo();
        expect(worksheet.getData()[0].slice(3)).to.eql(before);

        const [other] = jspreadsheet(root, {
            worksheets: [{ data: Array.from({ length: 20 }, (_, y) => [`${y + 1}`, y ? '' : '=SUM(A5:A12)', y ? '' : '=SUM($A$5:A$12)']) }],
        });
        other.deleteRow(11, 1);
        expect(other.getData()[0].slice(1)).to.eql(['=SUM(A5:A11)', '=SUM($A$5:A$11)']);
        other.deleteRow(9, 2);
        expect(other.getData()[0].slice(1)).to.eql(['=SUM(A5:A9)', '=SUM($A$5:A$9)']);
    });

    it('calculates again a formula whose range loses a cell from its middle', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=SUM(A2:A4)'],
                        ['2', ''],
                        ['3', ''],
                        ['4', ''],
                    ],
                },
            ],
            formulaEngine: (worksheet, expression) => `engine ${expression}`,
        });
        worksheet.deleteRow(2, 1);
        expect(worksheet.getValue('B1', true)).to.equal('engine =SUM(A2:A3)');
    });

    it('does not give #REF! to a formula that names a cell with the text #REF!', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['#REF!', '=A1&"x"']] }],
        });
        expect(worksheet.getValue('B1', true)).to.equal('#REF!x');
    });

    it('keeps the texts with updateReferences false', () => {
        const [worksheet] = jspreadsheet(root, {
            updateReferences: false,
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['5', '=A1'],
                        ['7', '=A2'],
                    ],
                },
            ],
        });

        worksheet.deleteRow(0, 1);
        expect(column(worksheet, 1)).to.eql(['=A1', '=A2']);
        worksheet.undo();
        expect(column(worksheet, 1)).to.eql(['', '=A1', '=A2']);
    });
});

describe('References after an insert, a delete and a move (#259)', () => {
    it('keeps string literals and the cells of other worksheets, and moves references with $ signs on an insert', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '="A1"&A1', '=$A$2+A$2+Sheet2!A2'],
                        ['2', '', ''],
                    ],
                },
            ],
        });
        const before = worksheet.getData()[0];

        worksheet.insertRow(1, 0, true);
        expect(worksheet.getData()[1]).to.eql(['1', '="A1"&A2', '=$A$3+A$3+Sheet2!A2']);
        expect(worksheet.getValue('B2', true)).to.equal('A11');
        worksheet.undo();
        expect(worksheet.getData()[0]).to.eql(before);

        worksheet.insertColumn(1, 0, true);
        expect(worksheet.getData()[0]).to.eql(['', '1', '="A1"&B1', '=$B$2+B$2+Sheet2!A2']);
        worksheet.undo();
        expect(worksheet.getData()[0]).to.eql(before);
    });

    it('moves references with $ signs on a delete, and the formula takes the new values of the cell', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1', '2', '=$B$1*10+B$1+$B1', '=SUM($B$1:$C$1)']] }],
        });

        worksheet.deleteColumn(0, 1);
        expect(worksheet.getData()[0]).to.eql(['2', '=$A$1*10+A$1+$A1', '=SUM($A$1:$B$1)']);
        worksheet.setValue('A1', '3');
        expect(worksheet.getValue('B1', true)).to.equal('36');

        worksheet.undo();
        worksheet.undo();
        expect(worksheet.getData()[0]).to.eql(['1', '2', '=$B$1*10+B$1+$B1', '=SUM($B$1:$C$1)']);
    });

    it('moves references with $ signs when a row moves', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '=$A$2'],
                        ['2', ''],
                    ],
                },
            ],
        });
        worksheet.moveRow(1, 0);
        expect(worksheet.getData()).to.eql([
            ['2', ''],
            ['1', '=$A$1'],
        ]);
        expect(worksheet.getValue('B2', true)).to.equal('2');
    });
});

describe('Footers after an insert or a delete (#260)', () => {
    const footers = (worksheet: Worksheet) => worksheet.options.footers!.map((row) => [...row]);
    const shown = (worksheet: Worksheet) =>
        Array.from(worksheet.tfoot!.children).map((tr) =>
            Array.from(tr.children)
                .slice(1, worksheet.headers.length + 1)
                .map((td) => td.textContent)
        );

    it('moves the references of a footer on an insert of rows, and shows the new result', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1'], ['2'], ['3']], footers: [['=SUM(A1:A3)']] }],
            formulaEngine: (worksheet, expression) => `engine ${expression}`,
        });

        worksheet.insertRow(1, 0, true);
        expect(footers(worksheet)).to.eql([['=SUM(A2:A4)']]);
        expect(shown(worksheet)).to.eql([['engine =SUM(A2:A4)']]);

        worksheet.undo();
        expect(footers(worksheet)).to.eql([['=SUM(A1:A3)']]);
        expect(shown(worksheet)).to.eql([['engine =SUM(A1:A3)']]);
    });

    it('gives #REF! to the references of a footer to deleted rows, and the undo gives back the text', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1'], ['2'], ['3']], footers: [['=A1+A3', '=SUM(A1:A3)']] }],
        });

        worksheet.deleteRow(0, 1);
        expect(footers(worksheet)).to.eql([['=#REF!+A2', '=SUM(A1:A2)']]);

        worksheet.undo();
        expect(footers(worksheet)).to.eql([['=A1+A3', '=SUM(A1:A3)']]);
        expect(shown(worksheet)[0][0]).to.equal('4');

        worksheet.redo();
        expect(footers(worksheet)).to.eql([['=#REF!+A2', '=SUM(A1:A2)']]);
    });

    it('moves the references of the footers of columns on an insert and a delete of columns', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1', '2', '3']], footers: [['=B1', '=A1+B1', '=C1']] }],
        });

        worksheet.insertColumn(1, 0, true);
        expect(footers(worksheet)).to.eql([['', '=C1', '=B1+C1', '=D1']]);
        worksheet.undo();
        expect(footers(worksheet)).to.eql([['=B1', '=A1+B1', '=C1']]);

        worksheet.deleteColumn(0, 1);
        expect(footers(worksheet)).to.eql([['=#REF!+A1', '=B1']]);
        expect(shown(worksheet)).to.eql([['#REF!', '3']]);

        worksheet.undo();
        expect(footers(worksheet)).to.eql([['=B1', '=A1+B1', '=C1']]);
        expect(shown(worksheet)).to.eql([['2', '3', '3']]);

        worksheet.redo();
        expect(footers(worksheet)).to.eql([['=#REF!+A1', '=B1']]);
    });

    it('removes the footer cells of the deleted columns (#262)', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1', '2', '3']], footers: [['a', 'b', 'c']] }],
        });
        const cells = () => Array.from(worksheet.tfoot!.children[0].children).map((td) => td.textContent);

        worksheet.deleteColumn(0, 1);
        expect(cells()).to.eql(['', 'b', 'c']);
        worksheet.undo();
        expect(cells()).to.eql(['', 'a', 'b', 'c']);
        worksheet.insertColumn(1, 0, true);
        worksheet.undo();
        expect(cells()).to.eql(['', 'a', 'b', 'c']);
    });
});
