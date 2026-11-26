import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

describe('Use footers', () => {
    it('Start the worksheet with a footer', () => {
        jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    freezeColumns: 2,
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    footers: [
                        ['a', 'b', 'c'],
                        // TYPING: footers is typed string[][], but the library also shows numbers
                        [1, 2, 3] as unknown as string[],
                    ],
                },
            ],
        });

        const footerTag = root.querySelector('tfoot')!;

        const firstRow = footerTag.children[0];

        expect(firstRow.children[1].innerHTML).to.equal('a');
        expect(firstRow.children[2].innerHTML).to.equal('b');
        expect(firstRow.children[3].innerHTML).to.equal('c');

        const secondRow = footerTag.children[1];

        expect(secondRow.children[1].innerHTML).to.equal('1');
        expect(secondRow.children[2].innerHTML).to.equal('2');
        expect(secondRow.children[3].innerHTML).to.equal('3');
    });

    it('insertColumn adds an empty footer cell for each new column', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['1', '2', '3']],
                    footers: [['a', 'b', 'c']],
                },
            ],
        })[0];

        const footerTexts = () => Array.prototype.slice.call(root.querySelector('tfoot')!.children[0].children, 1).map((td: Element) => td.textContent);

        worksheet.insertColumn(2, 1, false);

        expect(worksheet.options.footers![0]).to.eql(['a', 'b', '', '', 'c']);
        expect(footerTexts()).to.eql(['a', 'b', '', '', 'c']);

        worksheet.undo();

        expect(worksheet.options.footers![0]).to.eql(['a', 'b', 'c']);

        worksheet.redo();

        expect(worksheet.options.footers![0]).to.eql(['a', 'b', '', '', 'c']);
    });

    it('runs a footer formula without coordinates, also when it refers to the data cell at its position (issue #18)', () => {
        const calls: [string, number | undefined, number | undefined][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['1', '2', '=A1'],
                        ['3', '4', '5'],
                    ],
                    footers: [['=C1+C2', '=B1+B2', '=A1+A2']],
                },
            ],
            onbeforeformula: (instance, expression, x, y) => {
                calls.push([expression, x, y]);
            },
        })[0];

        expect(calls.filter((call) => call[1] === undefined)).to.eql([
            ['=C1+C2', undefined, undefined],
            ['=B1+B2', undefined, undefined],
            ['=A1+A2', undefined, undefined],
        ]);
        expect([...worksheet.tfoot!.children[0].children].slice(1).map((td) => td.textContent)).to.eql(['6', '6', '4']);
        // Only the cell C1 depends on A1
        expect(worksheet.formulas.chain).to.eql({ A1: ['C1'] });

        // The footers are calculated again after a change
        worksheet.setValue('A2', '5');

        expect(worksheet.tfoot!.children[0].children[3].textContent).to.equal('6');
    });
});
