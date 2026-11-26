import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('Create the options from a table', () => {
    function createTable(html: string) {
        const table = document.createElement('table');
        table.innerHTML = html;
        return table;
    }

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

        expect(options.columns!.map((column) => column.title)).to.eql(['Name', 'Age']);
        expect(options.data).to.eql([['x', '1'], ['y']]);
        expect(options.classes).to.eql({ A1: 'a' });
        expect(options.style).to.eql({ A1: 'color:red' });
        expect(options.mergeCells).to.eql({ A2: [2, 1] });
        expect(options.rows).to.eql({ 1: { height: '50px' } });
    });
});

describe('Create the nested headers from a table', () => {
    it('createFromTable gives the colspan of the nested headers as a number (issue #53)', () => {
        const table = document.createElement('table');
        table.innerHTML =
            '<thead><tr><th colspan="2">Group</th><th>Single</th></tr><tr><th>A</th><th>B</th><th>C</th></tr></thead><tbody><tr><td>1</td><td>2</td><td>3</td></tr></tbody>';

        const options = jspreadsheet.helpers.createFromTable(table, {});

        expect(options.nestedHeaders![0][0].colspan).to.equal(2);
        expect(options.nestedHeaders![0][1].colspan).to.equal(1);
    });
});

describe('Read the options of a table', () => {
    const createTable = (html: string) => {
        const table = document.createElement('table');
        table.innerHTML = html;
        return table;
    };

    it('logs a message for an element that is not a table', () => {
        const log = console.log;
        const messages: unknown[] = [];
        console.log = (message: unknown) => messages.push(message);

        try {
            // @ts-expect-error -- the test passes an element that is not a table, without options
            expect(jspreadsheet.helpers.createFromTable(document.createElement('div'))).to.equal(undefined);
        } finally {
            console.log = log;
        }

        expect(messages).to.eql(['Element is not a table']);
    });

    it('reads the widths of the colgroup and the options of the headers', () => {
        const table = createTable(
            '<colgroup><col style="width: 120px"><col width="80"><col></colgroup>' +
                '<thead><tr><th data-celltype="numeric" name="amount" id="a" data-mask="0.00" style="text-align: right">Amount</th><th>Name</th><th>Other</th></tr></thead>' +
                '<tbody><tr><td>1</td><td>x</td><td>y</td></tr></tbody>'
        );

        // TYPING: the options of createFromTable are optional, but the typings require them
        // @ts-expect-error -- the test reads a table without options
        const options = jspreadsheet.helpers.createFromTable(table);

        expect(options.columns![0]).to.eql({ width: '50px', type: 'numeric', title: 'Amount', align: 'right', name: 'amount', id: 'a', mask: '0.00' });
        expect(options.columns![1]).to.eql({ width: '50px', type: 'text', title: 'Name' });
    });

    it('reads the formulas, the bold cells, the hidden cells and the footers', () => {
        const table = createTable(
            '<tbody><tr><td data-formula="A2*2"></td><td data-formula="=B2"></td><td class="styleBold">b</td><td class="styleBold" style="color: red">c</td><td style="display: none">d</td></tr></tbody>' +
                '<tfoot><tr><td>Total</td><td>2</td></tr></tfoot>'
        );

        // TYPING: the options of createFromTable are optional, but the typings require them
        // @ts-expect-error -- the test reads a table without options
        const options = jspreadsheet.helpers.createFromTable(table);

        expect(options.data).to.eql([['=A2*2', '=B2', 'b', 'c', 'd']]);
        // A hidden cell is shown, so it has no style left
        expect(options.style).to.eql({ C1: 'font-weight:bold;', D1: 'color: red; font-weight:bold;' });
        expect(options.footers).to.eql([['Total', '2']]);
    });

    it('guesses the calendar and dropdown columns', () => {
        let html = '';
        for (let i = 0; i < 20; i++) {
            html += `<tr><td>2024-01-${String(i + 1).padStart(2, '0')}</td><td>${i % 2 ? 'yes' : 'no'}</td><td>${i}</td></tr>`;
        }
        const table = createTable(`<thead><tr><th>Date</th><th>Answer</th><th>Number</th></tr></thead><tbody>${html}</tbody>`);

        const options = jspreadsheet.helpers.createFromTable(table, { parseTableAutoCellType: true });

        expect(options.columns!.map((column) => column.type)).to.eql(['calendar', 'dropdown', 'text']);
        expect((options.columns![1] as Jspreadsheet.DropdownColumn).source).to.eql(['no', 'yes']);
    });
});

describe('Parse a CSV text', () => {
    it('reads the quoted fields, the quotes in them and each kind of line break', () => {
        expect(jspreadsheet.helpers.parseCSV('a,"b,c","d ""e"""\r\nf\rg;h\ni')).to.eql([
            ['a', 'b,c', 'd "e"'],
            ['f', '', ''],
            ['g;h', '', ''],
            ['i', '', ''],
        ]);
    });

    it('splits the fields at the delimiter', () => {
        expect(jspreadsheet.helpers.parseCSV('a;b\nc;d', ';')).to.eql([
            ['a', 'b'],
            ['c', 'd'],
        ]);
    });
});

describe('Invert an object', () => {
    it('makes the values the keys', () => {
        const inverted = jspreadsheet.helpers.invert({ a: 1, b: 'x' });

        expect(inverted[1]).to.equal('a');
        expect(inverted.x).to.equal('b');
    });
});

describe('Show a value with the mask of its column', () => {
    it('uses the mask, the format or the locale of the column', async () => {
        const worksheet = await new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [[1.5, 1.5, 3, 1234.5]],
                        columns: [
                            { type: 'numeric', mask: '0.00' },
                            { type: 'numeric', format: '0.0' },
                            { type: 'numeric', locale: 'en-US', options: { style: 'currency', currency: 'USD' } },
                            { type: 'numeric', mask: '#.##0,00', decimal: ',' },
                        ],
                    },
                ],
                onload: () => resolve(instance[0]),
            });
        });

        expect(worksheet.records[0].map((record) => record.element.textContent)).to.eql(['1.50', '1.5', '$3', '1.234,50']);
    });
});
