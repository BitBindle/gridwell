import { expect } from 'chai';

import gridwell from './support/gridwell';

describe('Cells', () => {
    const create = () =>
        gridwell(root, {
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

describe('Render the cells', () => {
    it('hides the overflow of the cell on the left when textOverflow is on', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [['A long text that overflows', 'B', '']],
                    textOverflow: true,
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').style.overflow).to.equal('hidden');
        expect(worksheet.getCell('B1').style.overflow).to.equal('');
        expect(worksheet.getCell('C1').style.overflow).to.equal('hidden');
    });

    it('calculates the formula in capitals outside the quotes when secureFormulas is on', () => {
        const create = (secureFormulas: boolean) =>
            gridwell(root, {
                worksheets: [
                    {
                        data: [['ab', '=a1&"cd"']],
                        secureFormulas,
                    },
                ],
            })[0];

        expect(create(false).getCell('B1').textContent).to.equal('#ERROR');

        root.innerHTML = '';

        const worksheet = create(true);
        expect(worksheet.getCell('B1').textContent).to.equal('abcd');
        // The data keeps the formula as it was written
        expect(worksheet.getValue('B1')).to.equal('=a1&"cd"');
    });

    it('marks a negative number red when the mask asks for it', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [[-5], [5]],
                    columns: [{ mask: '0;[Red]0' }],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').classList.contains('red')).to.equal(true);
        expect(worksheet.getCell('A2').classList.contains('red')).to.equal(false);

        worksheet.setValue('A1', 3);

        expect(worksheet.getCell('A1').classList.contains('red')).to.equal(false);
    });

    it('removes the scripts from the HTML of a cell when parseHTML is on', () => {
        const worksheet = gridwell(root, {
            parseHTML: true,
            worksheets: [
                {
                    data: [['<b>Bold</b><script>window.x = 1</script>']],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').innerHTML).to.equal('<b>Bold</b>');

        worksheet.setValue('A1', '<i>Italic</i><script>window.x = 1</script>');

        expect(worksheet.getCell('A1').innerHTML).to.equal('<i>Italic</i>');
    });

    it('shows the text of the items of a dropdown column', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [['1;2'], ['3']],
                    columns: [
                        {
                            type: 'dropdown',
                            multiple: true,
                            source: [
                                { id: '1', name: 'One' },
                                { id: '2', name: 'Two' },
                                { id: '3', name: 'Three' },
                            ],
                        },
                    ],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').textContent).to.equal('One; Two');
        expect(worksheet.getCell('A2').textContent).to.equal('Three');

        worksheet.setValue('A2', '2');

        expect(worksheet.getCell('A2').textContent).to.equal('Two');
    });

    it('sets the value of a radio cell when it is clicked', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [[false], [true]],
                    columns: [{ type: 'radio' }],
                },
            ],
        })[0];

        const input = worksheet.getCell('A1').querySelector('input')!;
        expect(input.type).to.equal('radio');
        expect(input.name).to.equal('c0');

        input.click();

        expect(worksheet.getValue('A1')).to.equal(true);
    });

    it('shows a square of the color for a color column with the render square', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [['#ff0000']],
                    columns: [{ type: 'color', render: 'square' }],
                },
            ],
        })[0];

        const square = worksheet.getCell('A1').querySelector<HTMLElement>('div.color')!;
        expect(square.style.backgroundColor).to.equal('rgb(255, 0, 0)');

        worksheet.setValue('A1', '#00ff00');

        expect(worksheet.getCell('A1').querySelector<HTMLElement>('div.color')!.style.backgroundColor).to.equal('rgb(0, 255, 0)');
    });
});

describe('Update the table after the cells change', () => {
    it('adds the spare rows and columns that minSpareRows and minSpareCols ask for', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [
                        ['A', 'B'],
                        ['C', 'D'],
                    ],
                    minSpareRows: 1,
                    minSpareCols: 1,
                },
            ],
        })[0];

        expect(worksheet.rows.length).to.equal(3);
        expect(worksheet.headers.length).to.equal(3);

        worksheet.setValue('C3', 'E');

        expect(worksheet.rows.length).to.equal(4);
        expect(worksheet.headers.length).to.equal(4);
    });
});
