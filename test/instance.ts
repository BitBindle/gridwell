import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

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

        const table = root.querySelector('tbody')!;
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

        const table = root.querySelector('tbody')!;
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

        const headerWorksheets = root.querySelector('.jtabs-headers')!;

        expect(headerWorksheets.children[0].innerHTML).to.include('Countries');
        expect(headerWorksheets.children[0].innerHTML).not.to.include('Employees');
        expect(headerWorksheets.children[1].innerHTML).to.include('Employees');
        expect(headerWorksheets.children[1].innerHTML).not.to.include('Countries');
    });
});

describe('Create a jspreadsheet instance from a table', () => {
    const createTable = () => {
        const table = document.createElement('table');
        table.innerHTML = '<tbody><tr><td>a</td><td>b</td></tr></tbody>';
        return table;
    };

    it('throws when the table has no parent (issue #55)', () => {
        expect(() => jspreadsheet(createTable(), { worksheets: [] })).to.throw('the table must be in a parent element');
    });

    it('creates the worksheet of a table without options or without worksheets', async () => {
        const first = root.appendChild(createTable());
        // The second table is replaced by a div in its own parent
        const parent = document.body.appendChild(document.createElement('div'));
        const second = parent.appendChild(createTable());

        try {
            const firstWorksheets = jspreadsheet(first);
            const secondWorksheets = jspreadsheet(second, {});
            await new Promise((resolve) => setTimeout(resolve));

            expect(firstWorksheets[0].getData()).to.eql([['a', 'b']]);
            expect(secondWorksheets[0].getData()).to.eql([['a', 'b']]);
        } finally {
            // The table was replaced by the div of the spreadsheet
            jspreadsheet.destroy(parent.firstChild as Jspreadsheet.JspreadsheetInstanceElement);
            parent.remove();
        }
    });

    it('replaces a table in a parent with the spreadsheet', async () => {
        const table = root.appendChild(createTable());

        const worksheets = jspreadsheet(table, { worksheets: [] });
        await new Promise((resolve) => setTimeout(resolve));

        expect(table.parentNode).to.equal(null);
        expect(worksheets.length).to.equal(1);
        expect(worksheets[0].getData()).to.eql([['a', 'b']]);
    });
});

describe('Use getWorksheetInstanceByName', () => {
    // The spreadsheet joins jspreadsheet.spreadsheet after a promise resolves
    const create = async (namespace: string) => {
        const worksheets = jspreadsheet(root, {
            namespace,
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
            // Without a name, the worksheets of the namespace are returned by name
            const result = jspreadsheet.getWorksheetInstanceByName(worksheetName, 'allByName') as Record<string, Jspreadsheet.WorksheetInstance>;

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
        range.setStart(element.firstChild!, 5);
        range.collapse(true);
        const selection = window.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);

        expect(jspreadsheet.helpers.getCaretIndex(element)).to.equal(5);
    });

    it('getCaretIndex reads the selection of the window for a spreadsheet whose root has no getSelection (issue #46)', async () => {
        const host = document.body.appendChild(document.createElement('div'));

        try {
            const worksheets = jspreadsheet(root, { root: host, worksheets: [{ data: [['Hello world']] }] });
            await new Promise((resolve) => setTimeout(resolve));

            const cell = worksheets[0].getCell('A1');
            const range = document.createRange();
            range.setStart(cell.firstChild!, 5);
            range.collapse(true);
            window.getSelection()!.removeAllRanges();
            window.getSelection()!.addRange(range);

            expect(jspreadsheet.helpers.getCaretIndex.call(worksheets[0].parent, cell)).to.equal(5);
        } finally {
            host.remove();
        }
    });
});

describe('Worksheet and spreadsheet methods', () => {
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet[]>((resolve) => {
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

        // @ts-expect-error -- the replacement returns a string, to tell it from the method of the class
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

        const filter = root.querySelector('.jss_filter')!;
        expect(filter).to.not.equal(null);
        expect(filter.querySelector('label')!.textContent).to.equal('Search: ');
        expect(filter.querySelector('.jss_search')).to.equal(worksheet.searchInput);

        const dropdown = filter.querySelector<HTMLSelectElement>('.jss_pagination_dropdown')!;
        expect(Array.from(dropdown.options).map((option) => option.value)).to.eql(['2', '5']);
        expect(dropdown.value).to.equal('2');
        expect(filter.children[0].textContent).to.equal('Show 25 entries');

        dropdown.value = '5';
        // The library sets a handler that takes no event
        (dropdown.onchange as () => void)();
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

    it('read the boolean options strictly, so 1 is not true and 0 is not false', () => {
        // @ts-expect-error -- the test passes numbers, to check that the boolean options are read strictly
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2], search: 1, editable: 0 }] })[0];

        expect(root.querySelector('.jss_filter')).to.equal(null);

        worksheet.openEditor(worksheet.getCell('A1'));
        expect(worksheet.editor.isOpen()).to.equal(true);
    });

    it('build a row of filters, with the hidden columns hidden', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 2], filters: true, columns: [{}, { type: 'hidden' }, {}] }],
        })[0];

        const cells = worksheet.filter!.querySelectorAll<HTMLElement>('.jss_column_filter');
        expect(cells.length).to.equal(3);
        expect(Array.from(cells).map((cell) => cell.getAttribute('data-x'))).to.eql(['0', '1', '2']);
        expect(cells[1].style.display).to.equal('none');
        expect(worksheet.thead.lastElementChild).to.equal(worksheet.filter);
    });

    it('put the elements of the worksheet in order', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        expect(Array.from(worksheet.element.children).map((element) => element.className)).to.eql(['jss_content', 'jss_pagination', 'jss_about']);
        expect(Array.from(worksheet.content.children)).to.eql([worksheet.zoomed]);
        expect(Array.from(worksheet.zoomed.children)).to.eql([worksheet.table, worksheet.corner, worksheet.clipboard.textarea]);
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

    it('stops a text selection on the corner without an inline event handler', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        // A Content Security Policy without 'unsafe-inline' blocks an inline handler and logs an error
        expect(worksheet.corner.hasAttribute('onselectstart')).to.equal(false);

        const event = new Event('selectstart', { bubbles: true, cancelable: true });
        worksheet.corner.dispatchEvent(event);
        expect(event.defaultPrevented).to.equal(true);
    });

    it('set the classes of the text overflow, the corner, the drag and the resize', () => {
        let worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] })[0];

        expect(worksheet.table.classList.contains('jss_overflow')).to.equal(true);
        expect(worksheet.corner.style.display).to.equal('');
        expect(worksheet.thead.className).to.equal('draggable resizable');
        expect(worksheet.tbody.className).to.equal('draggable resizable');

        // root has the spreadsheet
        jspreadsheet.destroy(root as Jspreadsheet.JspreadsheetInstanceElement);
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

describe('Destroy a spreadsheet', () => {
    it('clears the active worksheet when it belongs to the destroyed spreadsheet, and keeps it otherwise', async () => {
        const elements = [0, 1].map(() => document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement));

        try {
            jspreadsheet(elements[0], { worksheets: [{ minDimensions: [1, 1] }] });
            await new Promise((resolve) => setTimeout(resolve));
            const [second] = jspreadsheet(elements[1], { worksheets: [{ minDimensions: [1, 1] }] });
            await new Promise((resolve) => setTimeout(resolve));

            // The active worksheet keeps the destroyed spreadsheet in memory
            expect(jspreadsheet.current).to.equal(second);
            jspreadsheet.destroy(elements[0]);
            expect(jspreadsheet.current).to.equal(second);

            jspreadsheet.destroy(elements[1]);
            expect(jspreadsheet.current).to.equal(null);
        } finally {
            for (const element of elements) {
                element.remove();
            }
        }
    });

    it('clears the active worksheet when the application set a worksheet that deleteWorksheet removed', async () => {
        const element = document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement);

        try {
            jspreadsheet(element, { worksheets: [{ minDimensions: [1, 1] }, { minDimensions: [1, 1] }] });
            await new Promise((resolve) => setTimeout(resolve));
            element.spreadsheet.worksheets[1].openWorksheet(1);
            const removed = element.spreadsheet.worksheets[1];
            expect(jspreadsheet.current).to.equal(removed);
            removed.deleteWorksheet(1);
            await new Promise((resolve) => setTimeout(resolve));

            // The removed worksheet is not in the list, but it keeps the spreadsheet as its parent. deleteWorksheet
            // clears the active worksheet, so the application sets it here.
            expect(element.spreadsheet.worksheets).to.not.include(removed);
            jspreadsheet.current = removed;
            jspreadsheet.destroy(element);
            expect(jspreadsheet.current).to.equal(null);
        } finally {
            element.remove();
        }
    });

    it('stops the timer of a search when it destroys the spreadsheet of the active worksheet', async () => {
        const element = document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement);

        try {
            const [worksheet] = jspreadsheet(element, { worksheets: [{ minDimensions: [1, 1], search: true }] });
            await new Promise((resolve) => setTimeout(resolve));
            const input = element.querySelector('.jss_search') as HTMLInputElement;
            input.value = 'a';
            input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'a' }));
            expect(jspreadsheet.current).to.equal(worksheet);

            jspreadsheet.destroy(element);

            // The search timer runs 200 ms after the key
            await new Promise((resolve) => setTimeout(resolve, 250));
            expect(jspreadsheet.current).to.equal(null);
        } finally {
            element.remove();
        }
    });
});

describe('Delete the active worksheet (issue #182)', () => {
    it('sends no key to the worksheet that deleteWorksheet removed', async () => {
        const element = document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement);

        try {
            const worksheets = jspreadsheet(element, {
                worksheets: [
                    { minDimensions: [3, 3], worksheetName: 'A' },
                    { minDimensions: [3, 3], worksheetName: 'B' },
                ],
            });
            await new Promise((resolve) => setTimeout(resolve));
            worksheets[0].openWorksheet(1);
            const removed = worksheets[1];
            removed.updateSelectionFromCoords(0, 0, 0, 0);
            expect(jspreadsheet.current).to.equal(removed);

            removed.deleteWorksheet(1);
            expect(jspreadsheet.current).to.equal(null);
            document.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowDown' }));

            expect(removed.selection.range!.map(Number)).to.eql([0, 0, 0, 0]);
            expect(worksheets[0].selection.range).to.equal(null);
        } finally {
            jspreadsheet.destroy(element);
            element.remove();
        }
    });

    it('keeps the active worksheet when deleteWorksheet removes another worksheet', async () => {
        const element = document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement);

        try {
            const worksheets = jspreadsheet(element, { worksheets: [{ minDimensions: [1, 1] }, { minDimensions: [1, 1] }] });
            await new Promise((resolve) => setTimeout(resolve));
            jspreadsheet.current = worksheets[0];

            worksheets[0].deleteWorksheet(1);

            expect(jspreadsheet.current).to.equal(worksheets[0]);
        } finally {
            jspreadsheet.destroy(element);
            element.remove();
        }
    });

    it('stops the timer of a search when deleteWorksheet removes the active worksheet', async () => {
        const element = document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement);

        try {
            const worksheets = jspreadsheet(element, { worksheets: [{ minDimensions: [1, 1], search: true }, { minDimensions: [1, 1] }] });
            await new Promise((resolve) => setTimeout(resolve));
            jspreadsheet.current = worksheets[0];
            const input = element.querySelector('.jss_search') as HTMLInputElement;
            input.value = 'a';
            input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'a' }));
            expect(jspreadsheet.current).to.equal(worksheets[0]);

            worksheets[0].deleteWorksheet(0);

            expect((jspreadsheet as unknown as { timeControl: unknown }).timeControl).to.equal(null);
            // The search timer runs 200 ms after the key
            await new Promise((resolve) => setTimeout(resolve, 250));
            expect(jspreadsheet.current).to.equal(null);
        } finally {
            jspreadsheet.destroy(element);
            element.remove();
        }
    });
});

describe('Destroy all spreadsheets', () => {
    it('destroys each spreadsheet (issue #114)', async () => {
        // Each element gets its spreadsheet
        const elements = [0, 1, 2].map(() => document.body.appendChild(document.createElement('div') as Jspreadsheet.JspreadsheetInstanceElement));

        try {
            for (const element of elements) {
                jspreadsheet(element, { worksheets: [{ minDimensions: [1, 1] }] });
            }
            await new Promise((resolve) => setTimeout(resolve));
            const spreadsheets = elements.map((element) => element.spreadsheet);

            jspreadsheet.destroyAll();

            // The other tests reuse root, so the list can keep spreadsheets whose element has a newer one
            expect(spreadsheets.filter((spreadsheet) => jspreadsheet.spreadsheet.includes(spreadsheet))).to.eql([]);
            expect(elements.map((element) => Boolean(element.spreadsheet))).to.eql([false, false, false]);
            expect(elements.map((element) => element.innerHTML)).to.eql(['', '', '']);
        } finally {
            for (const element of elements) {
                element.remove();
            }
        }
    });
});
