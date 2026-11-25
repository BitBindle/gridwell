const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Cells', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['US', 'Apples'],
                        ['UK', 'Carrots'],
                    ],
                },
            ],
        })[0];

    it('Get a cell by name or by coordinates', () => {
        const worksheet = create();

        const cell = worksheet.getCell('B2');

        expect(cell.tagName).to.equal('TD');
        expect(cell.getAttribute('data-x')).to.equal('1');
        expect(cell.getAttribute('data-y')).to.equal('1');
        expect(worksheet.getCell(1, 1)).to.equal(cell);
        expect(worksheet.getCellFromCoords(1, 1)).to.equal(cell);
    });

    it('Get the label of a cell', () => {
        const worksheet = create();

        expect(worksheet.getLabel('B2')).to.equal('Carrots');
        expect(worksheet.getLabel(0, 1)).to.equal('UK');
    });

    it('Make a cell read-only by name or by element', () => {
        const worksheet = create();

        worksheet.setReadOnly('A2', true);
        expect(worksheet.isReadOnly('A2')).to.equal(true);
        expect(worksheet.isReadOnly(0, 1)).to.equal(true);
        expect(worksheet.getCell('A2').classList.contains('readonly')).to.equal(true);

        worksheet.setReadOnly(worksheet.getCell('A2'), false);
        expect(worksheet.isReadOnly(0, 1)).to.equal(false);

        worksheet.setReadOnly(worksheet.getCell('B1'), true);
        expect(worksheet.isReadOnly('B1')).to.equal(true);
    });
});
