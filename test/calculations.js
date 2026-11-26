const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Calculations', () => {
    it('Testing formula chain', () => {
        let test = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', ''],
                        ['', ''],
                        ['', ''],
                        ['', ''],
                        ['', ''],
                    ],
                },
            ],
        });

        test[0].setValue('B5', '=B3+A1');
        test[0].setValue('B3', '=A1+1');
        test[0].setValue('A1', '2');

        expect(test[0].getValue('B5', true)).to.equal('5');
    });

    it('keeps the backslashes and line breaks of a referenced text cell', () => {
        const texts = ['temp\\', 'line one\nline two', 'a\\tb'];
        let test = jspreadsheet(root, {
            worksheets: [
                {
                    data: texts.map((text, i) => [text, '=A' + (i + 1)]),
                },
            ],
        });

        expect(test[0].getValue('B1', true)).to.equal(texts[0]);
        expect(test[0].getValue('B2', true)).to.equal(texts[1]);
        expect(test[0].getValue('B3', true)).to.equal(texts[2]);
    });

    it('keeps the quotes and backslashes of a referenced text cell', () => {
        const text = 'Say "hi" \\ bye';
        let test = jspreadsheet(root, {
            worksheets: [
                {
                    data: [[text, '=A1', '=B1']],
                },
            ],
        });

        expect(test[0].getValue('B1', true)).to.equal(text);
        expect(test[0].getValue('C1', true)).to.equal(text);
    });

    it('Keeps special characters inside string literals', () => {
        let test = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['="Total: $"&"100"', '="Smith & Co"', '="A1 <> $B$2 ^ 2!"', '="Say ""hi"" & bye"'],
                        ['10', '="Value: $"&A2', '=B1&" "&A2', '="a=b"&"!"'],
                    ],
                },
            ],
        });

        expect(test[0].getValue('A1', true)).to.equal('Total: $100');
        expect(test[0].getValue('B1', true)).to.equal('Smith &amp; Co');
        expect(test[0].getValue('C1', true)).to.equal('A1 &lt;&gt; $B$2 ^ 2!');
        expect(test[0].getValue('D1', true)).to.equal('Say "hi" &amp; bye');
        expect(test[0].getValue('B2', true)).to.equal('Value: $10');
        expect(test[0].getValue('C2', true)).to.equal('Smith &amp; Co 10');
        expect(test[0].getValue('D2', true)).to.equal('a=b!');
    });

    describe('Test updating formulas when adding new rows', () => {
        it('1', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '=SUM(A2:C2)'],
                            ['4', '5', '6', '=SUM(A2:C2)'],
                            ['7', '8', '9', '=SUM(A2:C2)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertRow(1, 1, true);

            expect(test[0].getValue('D1')).to.equal('=SUM(A3:C3)');
            expect(test[0].getValue('D2')).to.equal('');
            expect(test[0].getValue('D3')).to.equal('=SUM(A3:C3)');
            expect(test[0].getValue('D4')).to.equal('=SUM(A3:C3)');
        });

        it('2', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '=SUM(A2:C3)'],
                            ['4', '5', '6', '=SUM(A2:C3)'],
                            ['7', '8', '9', '=SUM(A2:C3)'],
                            ['10', '11', '12', '=SUM(A2:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertRow(1, 1, true);

            expect(test[0].getValue('D1')).to.equal('=SUM(A3:C4)');
            expect(test[0].getValue('D2')).to.equal('');
            expect(test[0].getValue('D3')).to.equal('=SUM(A3:C4)');
            expect(test[0].getValue('D4')).to.equal('=SUM(A3:C4)');
            expect(test[0].getValue('D5')).to.equal('=SUM(A3:C4)');
        });

        it('3', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '=SUM(A2:C3)'],
                            ['4', '5', '6', '=SUM(A2:C3)'],
                            ['7', '8', '9', '=SUM(A2:C3)'],
                            ['10', '11', '12', '=SUM(A2:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertRow(1, 1, false);

            expect(test[0].getValue('D1')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D2')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D3')).to.equal('');
            expect(test[0].getValue('D4')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D5')).to.equal('=SUM(A2:C4)');
        });

        it('4', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '=SUM(A2:C3)'],
                            ['4', '5', '6', '=SUM(A2:C3)'],
                            ['7', '8', '9', '=SUM(A2:C3)'],
                            ['10', '11', '12', '=SUM(A2:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertRow(1, 2, true);

            expect(test[0].getValue('D1')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D2')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D3')).to.equal('');
            expect(test[0].getValue('D4')).to.equal('=SUM(A2:C4)');
            expect(test[0].getValue('D5')).to.equal('=SUM(A2:C4)');
        });

        it('5', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '=SUM(A2:C3)'],
                            ['4', '5', '6', '=SUM(A2:C3)'],
                            ['7', '8', '9', '=SUM(A2:C3)'],
                            ['10', '11', '12', '=SUM(A2:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertRow(1, 2, false);

            expect(test[0].getValue('D1')).to.equal('=SUM(A2:C3)');
            expect(test[0].getValue('D2')).to.equal('=SUM(A2:C3)');
            expect(test[0].getValue('D3')).to.equal('=SUM(A2:C3)');
            expect(test[0].getValue('D4')).to.equal('');
            expect(test[0].getValue('D5')).to.equal('=SUM(A2:C3)');
        });
    });

    describe('Test updating formulas when adding new columns', () => {
        it('1', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3'],
                            ['4', '5', '6'],
                            ['7', '8', '9'],
                            ['=SUM(B1:B3)', '=SUM(B1:B3)', '=SUM(B1:B3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertColumn(1, 1, true);

            expect(test[0].getValue('A4')).to.equal('=SUM(C1:C3)');
            expect(test[0].getValue('B4')).to.equal('');
            expect(test[0].getValue('C4')).to.equal('=SUM(C1:C3)');
            expect(test[0].getValue('D4')).to.equal('=SUM(C1:C3)');
        });

        it('2', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '4'],
                            ['5', '6', '7', '8'],
                            ['9', '10', '11', '12'],
                            ['=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertColumn(1, 1, true);

            expect(test[0].getValue('A4')).to.equal('=SUM(C1:D3)');
            expect(test[0].getValue('B4')).to.equal('');
            expect(test[0].getValue('C4')).to.equal('=SUM(C1:D3)');
            expect(test[0].getValue('D4')).to.equal('=SUM(C1:D3)');
            expect(test[0].getValue('E4')).to.equal('=SUM(C1:D3)');
        });

        it('3', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '4'],
                            ['5', '6', '7', '8'],
                            ['9', '10', '11', '12'],
                            ['=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertColumn(1, 1, false);

            expect(test[0].getValue('A4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('B4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('C4')).to.equal('');
            expect(test[0].getValue('D4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('E4')).to.equal('=SUM(B1:D3)');
        });

        it('4', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '4'],
                            ['5', '6', '7', '8'],
                            ['9', '10', '11', '12'],
                            ['=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertColumn(1, 2, true);

            expect(test[0].getValue('A4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('B4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('C4')).to.equal('');
            expect(test[0].getValue('D4')).to.equal('=SUM(B1:D3)');
            expect(test[0].getValue('E4')).to.equal('=SUM(B1:D3)');
        });

        it('5', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['1', '2', '3', '4'],
                            ['5', '6', '7', '8'],
                            ['9', '10', '11', '12'],
                            ['=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)', '=SUM(B1:C3)'],
                        ],
                        worksheetName: 'sheet1',
                    },
                ],
            });

            test[0].insertColumn(1, 2, false);

            expect(test[0].getValue('A4')).to.equal('=SUM(B1:C3)');
            expect(test[0].getValue('B4')).to.equal('=SUM(B1:C3)');
            expect(test[0].getValue('C4')).to.equal('=SUM(B1:C3)');
            expect(test[0].getValue('D4')).to.equal('');
            expect(test[0].getValue('E4')).to.equal('=SUM(B1:C3)');
        });

        it('6', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [[1, 2, 3], [1, 2, 3], [1, 2, 3], [1, 2, 3], [1, 2, 3], ['=SUM(A1:A4)']],
                        minDimensions: [5, 5],
                    },
                ],
            });

            test[0].deleteRow(1);

            expect(test[0].getValue('A5')).to.equal('=SUM(A1:A3)');
        });
    });

    describe('formulaEngine', () => {
        it('calculates a formula with the engine of the application', () => {
            const calls = [];
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '2', '=A1+B1']],
                    },
                ],
                formulaEngine: (worksheet, expression, x, y) => {
                    calls.push([worksheet, expression, x, y]);
                    return 'engine ' + expression;
                },
            });

            expect(test[0].getValue('C1', true)).to.equal('engine =A1+B1');
            expect(calls.length).to.equal(1);
            expect(calls[0][0]).to.equal(test[0]);
            expect(calls[0].slice(1)).to.eql(['=A1+B1', 2, 0]);
        });

        it('gives the engine the formula text as typed', () => {
            const expressions = [];
            jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '2', '=SUM($A$1:B1)&"$A1:B2"']],
                    },
                ],
                formulaEngine: (worksheet, expression) => {
                    expressions.push(expression);
                    return '';
                },
            });

            expect(expressions).to.eql(['=SUM($A$1:B1)&"$A1:B2"']);
        });

        it('gives the engine the text that onbeforeformula returns', () => {
            const expressions = [];
            jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '=A1']],
                    },
                ],
                onbeforeformula: () => '=A1*2',
                formulaEngine: (worksheet, expression) => {
                    expressions.push(expression);
                    return '';
                },
            });

            expect(expressions).to.eql(['=A1*2']);
        });

        it('calculates the dependent cells again when a referenced cell changes', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '=DOUBLE(A1)']],
                    },
                ],
                formulaEngine: (worksheet) => Number(worksheet.getValueFromCoords(0, 0)) * 2,
            });

            expect(test[0].getValue('B1', true)).to.equal('2');

            test[0].setValue('A1', '5');

            expect(test[0].getValue('B1', true)).to.equal('10');
        });

        it('does not calculate the referenced cells before it calls the engine', () => {
            const expressions = [];
            jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['=1', '=A1']],
                    },
                ],
                formulaEngine: (worksheet, expression) => {
                    expressions.push(expression);
                    return '';
                },
            });

            expect(expressions).to.eql(['=1', '=A1']);
        });

        it('gives the engine the coordinates as numbers when a cell element is set', () => {
            const calls = [];
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '2', '']],
                    },
                ],
                formulaEngine: (worksheet, expression, x, y) => {
                    calls.push([expression, x, y]);
                    return '';
                },
            });

            test[0].setValue(test[0].getCell('C1'), '=A1*B1');

            expect(calls).to.eql([['=A1*B1', 2, 0]]);
        });

        it('gives the engine no coordinates when executeFormula gets none', () => {
            const calls = [];
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '2']],
                    },
                ],
                formulaEngine: (worksheet, expression, x, y) => {
                    calls.push([expression, x, y]);
                    return '';
                },
            });

            test[0].executeFormula('=A1+B1');

            expect(calls).to.eql([['=A1+B1', undefined, undefined]]);
        });

        it('shows #ERROR when the engine throws', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['=A2']],
                    },
                ],
                formulaEngine: () => {
                    throw new Error('Unknown function');
                },
            });

            expect(test[0].getValue('A1', true)).to.equal('#ERROR');
        });

        it('calculates with the formula plugin when the option is not set', () => {
            let test = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['1', '2', '=A1+B1']],
                    },
                ],
            });

            expect(test[0].getValue('C1', true)).to.equal('3');
        });
    });

    it('reads a formatted number with the decimal separator of its column, also in column 0', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1,5', '1,5', '=A1*2', '=B1*2']],
                    columns: [{ decimal: ',' }, { decimal: ',' }, {}, {}],
                },
            ],
        })[0];

        expect(worksheet.getValueFromCoords(2, 0, true)).to.equal('3');
        expect(worksheet.getValueFromCoords(3, 0, true)).to.equal('3');
    });
});

describe('The formula chain of a worksheet', () => {
    const create = (data) => jspreadsheet(root, { worksheets: [{ data }] })[0];

    it('is empty before a formula is calculated', () => {
        const worksheet = create([['1', '2']]);

        expect(worksheet.formulas.chain).to.eql({});
    });

    it('keeps the cells whose formulas refer to a cell, once each', () => {
        const worksheet = create([['1', '2', '=A1+B1', '=A1']]);

        worksheet.setValue('D1', '=A1*2');

        expect(worksheet.formulas.chain).to.eql({ A1: ['C1', 'D1'], B1: ['C1'] });
    });

    it('renames the cells after a row is added above them', () => {
        const worksheet = create([
            ['1', '=A1'],
            ['2', '=A2'],
        ]);

        worksheet.insertRow(1, 0, true);

        expect(worksheet.formulas.chain).to.eql({ A2: ['B2'], A3: ['B3'] });
        expect(worksheet.getValue('B2')).to.equal('=A2');
    });

    it('add and rename change the chain', () => {
        const worksheet = create([['1']]);

        worksheet.formulas.add('A1', 'B1');
        worksheet.formulas.add('A1', 'B1');
        worksheet.formulas.add('A1', 'C1');
        worksheet.formulas.rename({ A1: 'A2', C1: 'C2' });

        expect(worksheet.formulas.chain).to.eql({ A2: ['B1', 'C2'] });
    });

    it('calculates the cells in the chain again when a cell changes', () => {
        const worksheet = create([['1', '=A1+1']]);

        worksheet.setValue('A1', '5');

        expect(worksheet.getValue('B1', true)).to.equal('6');
        expect(worksheet.formulas.chain.A1).to.eql(['B1']);
    });

    it('is not shared between worksheets', async () => {
        const worksheets = jspreadsheet(root, { worksheets: [{ data: [['1', '=A1']] }, { data: [['2']] }] });
        await new Promise((resolve) => setTimeout(resolve));

        expect(worksheets[0].formulas).to.not.equal(worksheets[1].formulas);
        expect(worksheets[1].formulas.chain).to.eql({});
    });
});

describe('The formula chain', () => {
    it('keeps no formula of executeFormula without coordinates in the chain, so that setValue still works', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['1', '=A1']] }] })[0];

        expect(worksheet.executeFormula('=A1+1')).to.equal(2);
        expect(worksheet.formulas.chain).to.eql({ A1: ['B1'] });

        worksheet.setValue('A1', '7');

        expect(worksheet.getValue('B1', true)).to.equal('7');
    });

    it('removes a cell that is no longer a formula from the chain of the cell it referred to (issue #43)', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['1', '=A1*2', '=A1*3']] }] })[0];

        expect(worksheet.formulas.chain.A1).to.deep.equal(['B1', 'C1']);

        worksheet.setValue('B1', '5');
        worksheet.setValue('A1', '2');

        expect(worksheet.formulas.chain.A1).to.deep.equal(['C1']);
        expect(worksheet.getData()).to.deep.equal([['2', '5', '=A1*3']]);
        expect(worksheet.getValue('C1', true)).to.equal('6');
    });
});
