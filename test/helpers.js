const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Create the options from a table', () => {
    const createTable = function (html) {
        const table = document.createElement('table');
        table.innerHTML = html;
        return table;
    };

    it('reads the class, style, merge and height of each cell', () => {
        const table = createTable('<tr><td class="a" style="color:red">x</td><td>1</td></tr>' + '<tr style="height: 50px"><td colspan="2">y</td></tr>');

        const options = jspreadsheet.helpers.createFromTable(table, {});

        expect(options.data).to.eql([['x', '1'], ['y']]);
        expect(options.classes).to.eql({ A1: 'a' });
        expect(options.style).to.eql({ A1: 'color:red' });
        expect(options.mergeCells).to.eql({ A2: [2, 1] });
        expect(options.rows).to.eql({ 1: { height: '50px' } });
    });

    it('places the class, style, merge and height in the data rows when the first row is the header', () => {
        const table = createTable(
            '<tr><td>Name</td><td>Age</td></tr>' +
                '<tr><td class="a" style="color:red">x</td><td>1</td></tr>' +
                '<tr style="height: 50px"><td colspan="2">y</td></tr>'
        );

        const options = jspreadsheet.helpers.createFromTable(table, { parseTableFirstRowAsHeader: true });

        expect(options.columns.map((column) => column.title)).to.eql(['Name', 'Age']);
        expect(options.data).to.eql([['x', '1'], ['y']]);
        expect(options.classes).to.eql({ A1: 'a' });
        expect(options.style).to.eql({ A1: 'color:red' });
        expect(options.mergeCells).to.eql({ A2: [2, 1] });
        expect(options.rows).to.eql({ 1: { height: '50px' } });
    });
});
