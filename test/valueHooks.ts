import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

describe('The formatValue hook', () => {
    it('shows the text of the hook and keeps the value of the cell', () => {
        const calls: [number, number, unknown][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[45658, 'a']] }],
            formatValue: (instance, x, y, value) => {
                calls.push([x, y, value]);
                return x === 0 ? '2025-01-01' : undefined;
            },
        });

        expect(worksheet.getCell('A1').textContent).to.equal('2025-01-01');
        expect(worksheet.getCell('B1').textContent).to.equal('a');
        expect(worksheet.getValueFromCoords(0, 0)).to.equal(45658);
        expect(calls).to.deep.include([0, 0, 45658]);
        expect(calls).to.deep.include([1, 0, 'a']);
    });

    it('gets the result of a formula, not the formula', () => {
        const values: unknown[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['=1+1']] }],
            formulaEngine: () => 2,
            formatValue: (instance, x, y, value) => {
                values.push(value);
                return `${value} units`;
            },
        });

        expect(values).to.eql([2]);
        expect(worksheet.getCell('A1').textContent).to.equal('2 units');
    });

    it('formats a new value of setValue', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['1']] }],
            formatValue: (instance, x, y, value) => `[${value}]`,
        });

        worksheet.setValue('A1', 'b');

        expect(worksheet.getCell('A1').textContent).to.equal('[b]');
        expect(worksheet.getValue('A1')).to.equal('b');
    });

    it('keeps the mask of the column when the hook gives undefined', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[1234.5]], columns: [{ type: 'numeric', mask: '#,##0.00' }] }],
            formatValue: () => undefined,
        });

        expect(worksheet.getCell('A1').textContent).to.equal('1,234.50');
    });

    it('can set the alignment of the cell after the alignment of the column', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[1, 'a']], columns: [{ align: 'left' }, { align: 'left' }] }],
            formatValue: (instance, x, y, value, cell) => {
                if (cell && typeof value === 'number') {
                    cell.style.textAlign = 'right';
                }
                return undefined;
            },
        });

        expect(worksheet.getCell('A1').style.textAlign).to.equal('right');
        expect(worksheet.getCell('B1').style.textAlign).to.equal('left');
    });

    it('does not run for a column that renders its own control', () => {
        const columns: number[] = [];
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [[true, '1', '#ff0000', 'text']],
                    columns: [{ type: 'checkbox' }, { type: 'dropdown', source: ['1'] }, { type: 'color' }, { type: 'text' }],
                },
            ],
            formatValue: (instance, x) => {
                columns.push(x);
                return undefined;
            },
        });

        expect(columns).to.eql([3]);
    });
});

describe('The formatValue hook with other column types', () => {
    it('gives the cell to the hook when setValue changes an html column', () => {
        const option = global.Option;
        global.Option = window.Option;
        try {
            const cells: boolean[] = [];
            const [worksheet] = jspreadsheet(root, {
                worksheets: [{ data: [['a']], columns: [{ type: 'html' }] }],
                formatValue: (instance, x, y, value, cell) => {
                    cells.push(cell !== undefined);
                    return undefined;
                },
            });

            worksheet.setValue('A1', 'b');

            expect(cells).to.eql([true, true]);
        } finally {
            global.Option = option;
        }
    });

    it('does not run for a hidden column, also after setValue', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[1]], columns: [{ type: 'hidden' }] }],
            formatValue: () => 'text of the hook',
        });

        worksheet.setValue('A1', 5);

        expect(worksheet.getCell('A1').textContent).to.equal('5');
    });

    it('removes the red color of the mask from a cell that shows the text of the hook', () => {
        const hook: { text?: string } = {};
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[-5]], columns: [{ type: 'numeric', mask: '#,##0;[Red]-#,##0' }] }],
            formatValue: () => hook.text,
        });
        expect(worksheet.getCell('A1').classList.contains('red')).to.equal(true);

        hook.text = 'seven';
        worksheet.setValue('A1', 7);

        expect(worksheet.getCell('A1').textContent).to.equal('seven');
        expect(worksheet.getCell('A1').classList.contains('red')).to.equal(false);
    });

    it('keeps the alignment that the createCell of a custom type sets', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['a']],
                    columns: [
                        {
                            align: 'left',
                            type: {
                                createCell: (cell: HTMLTableCellElement) => {
                                    cell.style.textAlign = 'right';
                                    return cell;
                                },
                            },
                        },
                    ],
                },
            ],
        });

        expect(worksheet.getCell('A1').style.textAlign).to.equal('right');
    });
});

describe('The editValue hook', () => {
    it('shows the text of the hook in the editor, and the editor saves its text', () => {
        const calls: [number, number, Jspreadsheet.CellValue][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[45658]] }],
            editValue: (instance, x, y, value) => {
                calls.push([x, y, value]);
                return '2025-01-01';
            },
        });
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);
        const editor = cell.children[0] as HTMLInputElement;
        expect(editor.value).to.equal('2025-01-01');
        expect(calls).to.eql([[0, 0, 45658]]);

        editor.value = '2025-01-02';
        worksheet.closeEditor(cell, true);
        expect(worksheet.getValue('A1')).to.equal('2025-01-02');
    });

    it('does not run when the editor opens empty', () => {
        let calls = 0;
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            editValue: () => {
                calls++;
                return 'b';
            },
        });
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell, true);

        expect((cell.children[0] as HTMLInputElement).value).to.equal('');
        expect(calls).to.equal(0);
    });

    it('keeps the editor text of the mask when the hook gives undefined', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[1234.5]], columns: [{ type: 'numeric', mask: '$ #,##0.00' }] }],
            editValue: () => undefined,
        });
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);

        expect((cell.children[0] as HTMLInputElement).value).to.equal('1234.5');
    });

    it('keeps the value of the cell when the user leaves the text of the hook as it is', () => {
        const changes: unknown[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [[45658, 1234.5]], columns: [{}, { type: 'numeric', mask: '#,##0.00' }] }],
            editValue: (instance, x) => (x === 0 ? '2025-01-01' : '1.234,50'),
            onchange: (instance, cell, x, y, value) => {
                changes.push(value);
            },
        });

        for (const name of ['A1', 'B1']) {
            const cell = worksheet.getCell(name);
            worksheet.openEditor(cell);
            worksheet.closeEditor(cell, true);
        }

        expect(worksheet.getValueFromCoords(0, 0)).to.equal(45658);
        expect(worksheet.getValueFromCoords(1, 0)).to.equal(1234.5);
        expect(changes).to.eql([]);
    });

    it('shows an empty editor when the hook returns null', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            // @ts-expect-error -- a hook in JavaScript can return null
            editValue: () => null,
        });
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);

        expect((cell.children[0] as HTMLInputElement).value).to.equal('');
    });
});
