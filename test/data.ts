import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

describe('Use the data method', () => {
    it('getData and it returns the data properly', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [3, 2, 1],
                        [4, 5, 6],
                        [6, 5, 4],
                        [9, 12, 15],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const data = instance[0].getData();

        expect(data.length).to.eq(7);
        expect(data[0].length).to.eq(7);

        expect(data[0][0]).to.eq(1);
        expect(data[0][1]).to.eq(2);
        expect(data[0][2]).to.eq(3);
        expect(data[1][0]).to.eq(3);
        expect(data[1][1]).to.eq(2);
        expect(data[1][2]).to.eq(1);
        expect(data[2][0]).to.eq(4);
        expect(data[2][1]).to.eq(5);
        expect(data[2][2]).to.eq(6);
        expect(data[3][0]).to.eq(6);
        expect(data[3][1]).to.eq(5);
        expect(data[3][2]).to.eq(4);
        expect(data[4][0]).to.eq(9);
        expect(data[4][1]).to.eq(12);
        expect(data[4][2]).to.eq(15);
    });

    it('setData and it sets data properly', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    worksheetName: 'Countries',
                },
            ],
        });

        instance[0].setData([
            ['Hello', 'World'],
            ['Testing', 'CE'],
        ]);

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[1].innerHTML).not.to.include('World');
        expect(firstRow.children[2].innerHTML).to.include('World');
        expect(firstRow.children[2].innerHTML).not.to.include('Hello');

        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[1].innerHTML).not.to.include('CE');
        expect(secondRow.children[2].innerHTML).to.include('CE');
        expect(secondRow.children[2].innerHTML).not.to.include('Testing');
    });

    it('setValue and it sets the value of a cell', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');
        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('CE');

        instance[0].setValue('A1', 'New Value');

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].setValue('A1', 'olleH');
        instance[0].setValue('B1', 'dlroW');

        expect(firstRow.children[1].innerHTML).to.include('olleH');
        expect(firstRow.children[2].innerHTML).to.include('dlroW');

        instance[0].setValue('B2', 'TESTING');

        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('TESTING');
    });

    it('setValue with a list gives the value argument to each item without its own value', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [3, 1],
                },
            ],
        });

        instance[0].setValue([{ x: 0, y: 0, value: 'first' }, { x: 1, y: 0 }, 'C1'], 'default');

        expect(instance[0].getValue('A1')).to.equal('first');
        expect(instance[0].getValue('B1')).to.equal('default');
        expect(instance[0].getValue('C1')).to.equal('default');
    });

    it('getValue and it gets the value from the cell', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        expect(instance[0].getValue('A1')).to.include('Hello');
        expect(instance[0].getValue('B1')).to.include('World');
        expect(instance[0].getValue('A2')).to.include('Testing');
        expect(instance[0].getValue('B2')).to.include('CE');
    });

    it('getValueFromCoords and it gets the not processed cell value', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['=1+1', '=2+2'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        expect(instance[0].getValueFromCoords(0, 0)).to.include('=1+1');
        expect(instance[0].getValueFromCoords(1, 0)).to.include('=2+2');
        expect(instance[0].getValueFromCoords(0, 1)).to.include('Testing');
        expect(instance[0].getValueFromCoords(1, 1)).to.include('CE');
    });

    it('getValueFromCoords and it gets the processed cell value', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['=1+1', '=2+2'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        expect(instance[0].getValueFromCoords(0, 0, true)).to.include('2');
        expect(instance[0].getValueFromCoords(1, 0, true)).to.include('4');
        expect(instance[0].getValueFromCoords(0, 1, true)).to.include('Testing');
        expect(instance[0].getValueFromCoords(1, 1, true)).to.include('CE');
    });

    it('setValueFromCoords and it sets the value of a cell', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        expect(firstRow.children[1].innerHTML).to.include('Hello');
        expect(firstRow.children[2].innerHTML).to.include('World');
        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('CE');

        instance[0].setValueFromCoords(0, 0, 'New Value');

        expect(firstRow.children[1].innerHTML).to.include('New Value');
        expect(firstRow.children[2].innerHTML).to.include('World');

        instance[0].setValueFromCoords(0, 0, 'olleH');
        instance[0].setValueFromCoords(1, 0, 'dlroW');

        expect(firstRow.children[1].innerHTML).to.include('olleH');
        expect(firstRow.children[2].innerHTML).to.include('dlroW');

        instance[0].setValueFromCoords(1, 1, 'TESTING');

        expect(secondRow.children[1].innerHTML).to.include('Testing');
        expect(secondRow.children[2].innerHTML).to.include('TESTING');
    });

    it('setValueFromCoords history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        ['Hello', 'World'],
                        ['Testing', 'CE'],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;
        const firstRow = rows[0];

        instance[0].setValueFromCoords(0, 0, 'New Value');

        expect(firstRow.children[1].innerHTML).to.include('New Value');

        instance[0].undo();

        expect(firstRow.children[1].innerHTML).to.include('Hello');

        instance[0].redo();

        expect(firstRow.children[1].innerHTML).to.include('New Value');
    });
});

describe('Get the data as text, as objects or from a range', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        [1, 2, '=A1+B1'],
                        [4, 5, '=A2+B2'],
                    ],
                },
            ],
        })[0];

    it('getData with a delimiter returns the rows as text', () => {
        const worksheet = create();

        expect(worksheet.getData(false, false, ';')).to.equal('1;2;=A1+B1\r\n4;5;=A2+B2\r\n');
        expect(worksheet.getData(false, true, ',')).to.equal('1,2,3\r\n4,5,9\r\n');
    });

    it('getData with asJson returns each row as an object', () => {
        const worksheet = create();

        expect(worksheet.getData(false, false, null, true)).to.eql([
            { 0: 1, 1: 2, 2: '=A1+B1' },
            { 0: 4, 1: 5, 2: '=A2+B2' },
        ]);
    });

    it('getData with highlighted returns only the selected cells', () => {
        const worksheet = create();

        worksheet.updateSelectionFromCoords(1, 0, 2, 1);

        expect(worksheet.getData(true)).to.eql([
            [2, '=A1+B1'],
            [5, '=A2+B2'],
        ]);
        expect(worksheet.getData(true, true)).to.eql([
            ['2', '3'],
            ['5', '9'],
        ]);
    });

    it('getDataFromRange returns the values of a range', () => {
        const worksheet = create();

        expect(worksheet.getDataFromRange('B1:C2')).to.eql([
            [2, '=A1+B1'],
            [5, '=A2+B2'],
        ]);
        expect(worksheet.getDataFromRange('B1:C2', true)).to.eql([
            ['2', '3'],
            ['5', '9'],
        ]);
    });

    it('setData changes rows of objects to arrays by the column names', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    columns: [{ name: 'city' }, { name: 'country' }],
                },
            ],
        })[0];

        worksheet.setData([
            { country: 'France', city: 'Paris' },
            { city: 'Rome', country: 'Italy' },
        ]);

        expect(worksheet.getData()).to.eql([
            ['Paris', 'France'],
            ['Rome', 'Italy'],
        ]);
        expect(worksheet.getValue('A2')).to.equal('Rome');
    });
});

describe('Get the value of a cell that is missing or holds the text undefined', () => {
    it('getValue and getValueFromCoords return the text undefined of a cell (issue #36)', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['undefined', 'x']] }] })[0];

        expect(worksheet.getValue('A1')).to.equal('undefined');
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('undefined');
        expect(worksheet.getValue('B1')).to.equal('x');
    });

    it('getValue and getValueFromCoords return null for a column outside the data (issue #49)', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a', 'b'], ['c']], columns: [{}, {}] }] })[0];

        expect(worksheet.getValue('C1')).to.equal(null);
        expect(worksheet.getValueFromCoords(2, 0)).to.equal(null);
        expect(worksheet.getValue('A5')).to.equal(null);
        expect(worksheet.getValueFromCoords(0, 4)).to.equal(null);
    });
});
