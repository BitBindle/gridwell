const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Create a jspreadsheet instance', () => {
    it('and the dimensions are applied correctly', () => {
        jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [10, 10],
                    data: [[10, 20, 30]],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody');
        const rows = table.children;
        const firstRow = rows[0];

        // check that the amount of rows displayed is 10
        expect(rows.length).to.eq(10);

        // check that the amount of columns displayed is 10 + 1 (one from the row header)
        expect(firstRow.children.length).to.eq(10 + 1);
    });

    it('and the data is displayed correctly', () => {
        jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [10, 10],
                    data: [
                        [10, 20, 30],
                        [40, 50, 60],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        const table = root.querySelector('tbody');
        const rows = table.children;
        const firstRow = rows[0];
        const secondRow = rows[1];

        // check that [A1, B1, C1] received the data value
        expect(firstRow.children[1].innerHTML).to.include(10);
        expect(firstRow.children[1].innerHTML).not.to.include(20);
        expect(firstRow.children[1].innerHTML).not.to.include(30);
        expect(firstRow.children[2].innerHTML).to.include(20);
        expect(firstRow.children[2].innerHTML).not.to.include(10);
        expect(firstRow.children[2].innerHTML).not.to.include(30);
        expect(firstRow.children[3].innerHTML).to.include(30);
        expect(firstRow.children[3].innerHTML).not.to.include(10);
        expect(firstRow.children[3].innerHTML).not.to.include(20);

        // check that [A2, B2, C2] received the data value
        expect(secondRow.children[1].innerHTML).to.include(40);
        expect(secondRow.children[1].innerHTML).not.to.include(50);
        expect(secondRow.children[1].innerHTML).not.to.include(60);
        expect(secondRow.children[2].innerHTML).to.include(50);
        expect(secondRow.children[2].innerHTML).not.to.include(40);
        expect(secondRow.children[2].innerHTML).not.to.include(60);
        expect(secondRow.children[3].innerHTML).to.include(60);
        expect(secondRow.children[3].innerHTML).not.to.include(40);
        expect(secondRow.children[3].innerHTML).not.to.include(50);
    });

    it('and the worksheet names are displayed correctly', () => {
        jspreadsheet(root, {
            tabs: true,
            worksheets: [
                { minDimensions: [10, 10], worksheetName: 'Countries' },
                { minDimensions: [10, 10], worksheetName: 'Employees' },
            ],
        });

        const headerWorksheets = root.querySelector('.jtabs-headers');

        expect(headerWorksheets.children[0].innerHTML).to.include('Countries');
        expect(headerWorksheets.children[0].innerHTML).not.to.include('Employees');
        expect(headerWorksheets.children[1].innerHTML).to.include('Employees');
        expect(headerWorksheets.children[1].innerHTML).not.to.include('Countries');
    });
});

describe('Use getWorksheetInstanceByName', () => {
    // The spreadsheet joins jspreadsheet.spreadsheet after a promise resolves
    const create = async (namespace) => {
        const worksheets = jspreadsheet(root, {
            namespace: namespace,
            worksheets: [
                { minDimensions: [2, 2], worksheetName: 'Countries' },
                { minDimensions: [2, 2], worksheetName: 'Employees' },
            ],
        });

        await new Promise((resolve) => setTimeout(resolve));

        return worksheets;
    };

    it('returns the worksheet with the name in the namespace', async () => {
        const worksheets = await create('byName');

        expect(jspreadsheet.getWorksheetInstanceByName('Countries', 'byName')).to.equal(worksheets[0]);
        expect(jspreadsheet.getWorksheetInstanceByName('Employees', 'byName')).to.equal(worksheets[1]);
    });

    it('returns all worksheets of the namespace by name when the name is null or undefined', async () => {
        const worksheets = await create('allByName');

        for (const worksheetName of [null, undefined]) {
            const result = jspreadsheet.getWorksheetInstanceByName(worksheetName, 'allByName');

            expect(Object.keys(result)).to.deep.equal(['Countries', 'Employees']);
            expect(result.Countries).to.equal(worksheets[0]);
            expect(result.Employees).to.equal(worksheets[1]);
        }
    });

    it('returns undefined for a name that is not in the namespace', async () => {
        await create('unknownName');

        expect(jspreadsheet.getWorksheetInstanceByName('Products', 'unknownName')).to.equal(undefined);
    });

    it('returns an empty object for a namespace that does not exist', async () => {
        await create('existing');

        expect(jspreadsheet.getWorksheetInstanceByName('Countries', 'missing')).to.deep.equal({});
        expect(jspreadsheet.getWorksheetInstanceByName(null, 'missing')).to.deep.equal({});
    });
});

describe('Use the helpers', () => {
    it('getCaretIndex returns the position of the caret in an element', () => {
        const element = document.createElement('div');
        element.textContent = 'Hello world';
        root.appendChild(element);

        const range = document.createRange();
        range.setStart(element.firstChild, 5);
        range.collapse(true);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        expect(jspreadsheet.helpers.getCaretIndex(element)).to.equal(5);
    });
});

describe('Worksheet and spreadsheet methods', () => {
    const load = (options) =>
        new Promise((resolve) => {
            const worksheets = jspreadsheet(root, { ...options, onload: () => resolve(worksheets) });
        });

    it('shares the methods of the worksheets through the prototype', async () => {
        const [first, second] = await load({ worksheets: [{ data: [['a']] }, { data: [['b']] }] });

        expect(first.getData).to.be.a('function');
        expect(first.getData).to.equal(second.getData);
        expect(Object.keys(first)).to.not.include('getData');
        expect(first.getData()).to.deep.equal([['a']]);
        expect(second.getData()).to.deep.equal([['b']]);
    });

    it('shares the methods of the spreadsheets through the prototype', async () => {
        const [worksheet] = await load({ worksheets: [{ data: [['a']] }] });
        root.innerHTML = '';
        const [other] = await load({ worksheets: [{ data: [['b']] }] });

        expect(worksheet.parent.getConfig).to.be.a('function');
        expect(worksheet.parent.getConfig).to.equal(other.parent.getConfig);
        expect(Object.keys(worksheet.parent)).to.not.include('getConfig');
        expect(other.parent.getWorksheetActive()).to.equal(0);
    });

    it('lets a worksheet replace a method for itself', async () => {
        const [first, second] = await load({ worksheets: [{ data: [['a']] }, { data: [['b']] }] });

        first.getData = () => 'replaced';

        expect(first.getData()).to.equal('replaced');
        expect(second.getData()).to.deep.equal([['b']]);
    });
});

describe('The options of the worksheet table', () => {
    it('build the search, the number of rows on a page and the pagination', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 6], search: true, pagination: 2, paginationOptions: [2, 5] }],
        })[0];

        const filter = root.querySelector('.jss_filter');
        expect(filter).to.not.equal(null);
        expect(filter.querySelector('label').textContent).to.equal('Search: ');
        expect(filter.querySelector('.jss_search')).to.equal(worksheet.searchInput);

        const dropdown = filter.querySelector('.jss_pagination_dropdown');
        expect(Array.from(dropdown.options).map((option) => option.value)).to.eql(['2', '5']);
        expect(dropdown.value).to.equal('2');
        expect(filter.children[0].textContent).to.equal('Show 25entries');

        dropdown.value = '5';
        dropdown.onchange();
        expect(worksheet.options.pagination).to.equal(5);

        expect(worksheet.pagination.style.display).to.equal('');
        expect(worksheet.pagination.children.length).to.equal(2);
    });

    it('have no search and no pagination by default', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        expect(root.querySelector('.jss_filter')).to.equal(null);
        expect(root.querySelector('.jss_pagination_dropdown')).to.equal(null);
        expect(worksheet.pagination.style.display).to.equal('none');
    });

    it('build a row of filters, with the hidden columns hidden', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 2], filters: true, columns: [{}, { type: 'hidden' }, {}] }],
        })[0];

        const cells = worksheet.filter.querySelectorAll('.jss_column_filter');
        expect(cells.length).to.equal(3);
        expect(Array.from(cells).map((cell) => cell.getAttribute('data-x'))).to.eql(['0', '1', '2']);
        expect(cells[1].style.display).to.equal('none');
        expect(worksheet.thead.lastElementChild).to.equal(worksheet.filter);
    });

    it('put the elements of the worksheet in order', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        expect(Array.from(worksheet.element.children).map((element) => element.className)).to.eql(['jss_content', 'jss_pagination', 'jss_about']);
        expect(Array.from(worksheet.content.children)).to.eql([worksheet.table, worksheet.corner, worksheet.clipboard.textarea]);
        expect(Array.from(worksheet.table.children).map((element) => element.tagName)).to.eql(['COLGROUP', 'THEAD', 'TBODY']);
        expect(worksheet.table.querySelectorAll('colgroup col').length).to.equal(3);
        expect(worksheet.element.classList.contains('jss_container')).to.equal(true);
        expect(worksheet.element.jssWorksheet).to.equal(worksheet);
        expect(worksheet.element.jspreadsheet).to.equal(worksheet);
    });

    it('set the size and the overflow of the content', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2], tableOverflow: true, tableHeight: 300, tableWidth: '400px' }],
        })[0];

        expect(worksheet.content.style.overflowY).to.equal('auto');
        expect(worksheet.content.style.maxHeight).to.equal('300px');
        expect(worksheet.content.style.overflowX).to.equal('auto');
        expect(worksheet.content.style.width).to.equal('400px');
    });

    it('set the classes of the text overflow, the corner, the drag and the resize', () => {
        let worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        expect(worksheet.table.classList.contains('jss_overflow')).to.equal(true);
        expect(worksheet.corner.style.display).to.equal('');
        expect(worksheet.thead.className).to.equal('draggable resizable');
        expect(worksheet.tbody.className).to.equal('draggable resizable');

        jspreadsheet.destroy(root);
        worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [2, 2],
                    textOverflow: true,
                    selectionCopy: false,
                    columnDrag: false,
                    columnResize: false,
                    rowDrag: false,
                    rowResize: false,
                },
            ],
        })[0];

        expect(worksheet.table.classList.contains('jss_overflow')).to.equal(false);
        expect(worksheet.corner.style.display).to.equal('none');
        expect(worksheet.thead.className).to.equal('');
        expect(worksheet.tbody.className).to.equal('');
    });

    it('apply the style, the comments and the classes of the cells', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [2, 2],
                    style: { A1: 'color:red' },
                    comments: { B1: 'A note' },
                    classes: { A2: 'marked' },
                },
            ],
        })[0];

        expect(worksheet.getStyle('A1', 'color')).to.equal('red');
        expect(worksheet.options.style).to.eql(worksheet.getStyle());
        expect(worksheet.getComments('B1')).to.equal('A note');
        expect(worksheet.getCell('A2').classList.contains('marked')).to.equal(true);
    });
});
