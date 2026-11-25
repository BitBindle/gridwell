const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const createData = function (rows, columns) {
    const data = [];
    for (let j = 0; j < rows; j++) {
        const row = [];
        for (let i = 0; i < columns; i++) {
            row.push(j + '-' + i);
        }
        data.push(row);
    }
    return data;
};

// Returns the aria-rowindex of the row of each data cell and the aria-colindex of the cell, as 'row:column'
const getIndexes = function (worksheet) {
    return worksheet.records.map(function (row, y) {
        return row.map(function (record) {
            return worksheet.rows[y].element.getAttribute('aria-rowindex') + ':' + record.element.getAttribute('aria-colindex');
        });
    });
};

const getExpectedIndexes = function (rows, columns, headerRows) {
    const result = [];
    for (let j = 0; j < rows; j++) {
        const row = [];
        for (let i = 0; i < columns; i++) {
            row.push(j + 1 + (headerRows || 1) + ':' + (i + 2));
        }
        result.push(row);
    }
    return result;
};

describe('Grid roles for assistive technology', () => {
    it('Start the worksheet with the roles of the grid pattern', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 2] }],
        });
        const worksheet = instance[0];

        expect(worksheet.table.getAttribute('role')).to.equal('grid');
        expect(worksheet.thead.getAttribute('role')).to.equal('rowgroup');
        expect(worksheet.tbody.getAttribute('role')).to.equal('rowgroup');

        expect(worksheet.headerContainer.getAttribute('role')).to.equal('row');
        expect(worksheet.headerContainer.getAttribute('aria-rowindex')).to.equal('1');
        expect(worksheet.headerContainer.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.headers.map((td) => td.getAttribute('role'))).to.eql(['columnheader', 'columnheader', 'columnheader']);
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3', '4']);

        const row = worksheet.rows[1].element;
        expect(row.getAttribute('role')).to.equal('row');
        expect(row.getAttribute('aria-rowindex')).to.equal('3');
        expect(row.children[0].getAttribute('role')).to.equal('rowheader');
        expect(row.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.records[1][2].element.getAttribute('aria-colindex')).to.equal('4');

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('3');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');
    });

    it('Give the full size of the worksheet with lazy loading', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(300, 3), lazyLoading: true, tableOverflow: true, tableHeight: '300px' }],
        });
        const worksheet = instance[0];

        expect(worksheet.tbody.children.length).to.equal(100);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('301');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');

        // Go to the last page
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.last(false, true);

        const firstRow = worksheet.tbody.children[0];
        const lastRow = worksheet.tbody.children[worksheet.tbody.children.length - 1];
        expect(firstRow.getAttribute('data-y')).to.equal('200');
        expect(firstRow.getAttribute('aria-rowindex')).to.equal('202');
        expect(lastRow.getAttribute('aria-rowindex')).to.equal('301');
    });

    it('Give the full size of the worksheet with pagination', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(10, 2), pagination: 3 }],
        });
        const worksheet = instance[0];

        expect(worksheet.tbody.children.length).to.equal(3);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('11');

        worksheet.page(2);

        const rows = Array.prototype.map.call(worksheet.tbody.children, (tr) => tr.getAttribute('aria-rowindex'));
        expect(rows).to.eql(['8', '9', '10']);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('11');
    });

    it('Keep the row indexes after insertRow, deleteRow, moveRow and undo', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(4, 2) }],
        });
        const worksheet = instance[0];

        worksheet.insertRow(2, 0, true);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('7');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(6, 2));
        expect(worksheet.records[2][0].element.textContent).to.equal('0-0');

        worksheet.deleteRow(0, 3);
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('4');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 2));

        worksheet.moveRow(0, 2);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 2));

        worksheet.undo();
        worksheet.undo();
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('7');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(6, 2));
    });

    it('Keep the column indexes after insertColumn, deleteColumn and moveColumn', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 3) }],
        });
        const worksheet = instance[0];

        worksheet.insertColumn(2, 0, true);
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('6');
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3', '4', '5', '6']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 5));

        worksheet.deleteColumn(0, 3);
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('3');
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2));

        worksheet.moveColumn(0, 1);
        expect(worksheet.headers.map((td) => td.getAttribute('aria-colindex'))).to.eql(['2', '3']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2));
    });

    it('Keep the column indexes of merged cells after insertColumn', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 3), mergeCells: { B1: [2, 2] } }],
        });
        const worksheet = instance[0];

        worksheet.insertColumn(1, 0, true);

        expect(worksheet.getMerge('C1')).to.eql([2, 2]);
        expect(worksheet.records[0][2].element.getAttribute('aria-colindex')).to.equal('4');
        // The cells under the merged cell
        const cells = worksheet.options.mergeCells.C1[2];
        expect(cells.map((td) => td.getAttribute('data-x') + ':' + td.getAttribute('aria-colindex'))).to.eql(['3:5', '2:4', '3:5']);
    });

    it('Keep the row indexes after orderBy', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: [['c'], ['a'], ['b']] }],
        });
        const worksheet = instance[0];

        worksheet.orderBy(0, 0);

        expect(worksheet.getColumnData(0)).to.eql(['a', 'b', 'c']);
        expect(Array.prototype.map.call(worksheet.tbody.children, (tr) => tr.getAttribute('aria-rowindex'))).to.eql(['2', '3', '4']);
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(3, 1));
    });

    it('Keep the indexes of hidden rows and columns', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(3, 3) }],
        });
        const worksheet = instance[0];

        worksheet.hideRow(1);
        worksheet.hideColumn(1);

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('4');
        expect(worksheet.table.getAttribute('aria-colcount')).to.equal('4');
        expect(worksheet.rows[2].element.getAttribute('aria-rowindex')).to.equal('4');
        expect(worksheet.records[2][2].element.getAttribute('aria-colindex')).to.equal('4');
    });

    it('Count the rows of nested headers and footers', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: createData(2, 2),
                    nestedHeaders: [[{ title: 'Group', colspan: 2 }]],
                    footers: [['Total', '']],
                },
            ],
        });
        const worksheet = instance[0];
        const nested = worksheet.thead.children[0];
        const footer = worksheet.table.querySelector('tfoot > tr');

        expect(nested.getAttribute('role')).to.equal('row');
        expect(nested.getAttribute('aria-rowindex')).to.equal('1');
        expect(nested.children[1].getAttribute('role')).to.equal('columnheader');
        expect(worksheet.headerContainer.getAttribute('aria-rowindex')).to.equal('2');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(2, 2, 2));
        expect(footer.getAttribute('role')).to.equal('row');
        expect(footer.getAttribute('aria-rowindex')).to.equal('5');
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('5');

        worksheet.insertRow();

        expect(footer.getAttribute('aria-rowindex')).to.equal('6');
        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('6');
    });

    it('Keep the column indexes of nested header, filter and footer cells after hideColumn', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: createData(2, 4),
                    filters: true,
                    nestedHeaders: [
                        [
                            { title: 'G1', colspan: 2 },
                            { title: 'G2', colspan: 2 },
                        ],
                    ],
                    footers: [['w', 'x', 'y', 'z']],
                },
            ],
        })[0];
        const nested = worksheet.thead.children[0];
        const footer = worksheet.table.querySelector('tfoot > tr');

        worksheet.hideColumn(1);

        // The cells of each row in the column D have the aria-colindex of the header of the column D
        const expected = worksheet.headers[3].getAttribute('aria-colindex');
        expect(expected).to.equal('5');
        expect(worksheet.filter.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(footer.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(worksheet.records[0][3].element.getAttribute('aria-colindex')).to.equal(expected);
        // The nested header G2 starts in the column C
        expect(nested.children[2].getAttribute('aria-colindex')).to.equal('4');
        expect(nested.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.filter.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(footer.children[0].getAttribute('aria-colindex')).to.equal('1');

        worksheet.deleteColumn(0);

        expect(footer.children[3].getAttribute('aria-colindex')).to.equal(worksheet.headers[2].getAttribute('aria-colindex'));
        expect(footer.children[3].getAttribute('aria-colindex')).to.equal('4');
    });

    it('Update the size of the grid after setData', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: createData(2, 2) }],
        });
        const worksheet = instance[0];

        worksheet.setData(createData(5, 2));

        expect(worksheet.table.getAttribute('aria-rowcount')).to.equal('6');
        expect(getIndexes(worksheet)).to.eql(getExpectedIndexes(5, 2));
    });
});

describe('Accessible names', () => {
    it('Give English names without the labels option', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ worksheetName: 'Cars', minDimensions: [3, 3] }],
        });
        const worksheet = instance[0];
        expect(worksheet.table.getAttribute('aria-label')).to.equal('Cars');
        expect(worksheet.headerContainer.children[0].getAttribute('aria-label')).to.equal('Select all');
        expect(worksheet.clipboard.textarea.getAttribute('aria-label')).to.equal('Clipboard helper');

        worksheet.openEditor(worksheet.records[1][2].element);

        expect(worksheet.editor.cell.children[0].getAttribute('aria-label')).to.equal('Edit cell C2');
    });

    it('Use the texts of the labels option', () => {
        const instance = jspreadsheet(root, {
            labels: { selectAll: 'Alles auswählen', editCell: 'Zelle {0} bearbeiten', clipboardHelper: 'Zwischenablage' },
            worksheets: [
                {
                    minDimensions: [4, 3],
                    wordWrap: true,
                    columns: [{}, { type: 'dropdown', source: ['a', 'b'] }, {}, { type: 'calendar' }],
                },
            ],
        });
        const worksheet = instance[0];

        expect(worksheet.headerContainer.children[0].getAttribute('aria-label')).to.equal('Alles auswählen');
        expect(worksheet.clipboard.textarea.getAttribute('aria-label')).to.equal('Zwischenablage');

        worksheet.openEditor(worksheet.records[2][0].element);

        expect(worksheet.editor.cell.children[0].tagName).to.equal('TEXTAREA');
        expect(worksheet.editor.cell.children[0].getAttribute('aria-label')).to.equal('Zelle A3 bearbeiten');

        worksheet.closeEditor(worksheet.editor.cell, false);
        worksheet.openEditor(worksheet.records[0][1].element);

        expect(worksheet.editor.cell.querySelector('.jdropdown-header').getAttribute('aria-label')).to.equal('Zelle B1 bearbeiten');

        worksheet.closeEditor(worksheet.editor.cell, false);
        worksheet.openEditor(worksheet.records[2][3].element);

        expect(worksheet.editor.cell.children[0].getAttribute('aria-label')).to.equal('Zelle D3 bearbeiten');
    });

    it('Give the rich text editor a role and a name', () => {
        // The html cells use the Option constructor of the window, and the rich text editor uses execCommand, which
        // jsdom does not have
        const option = global.Option;
        const execCommand = document.execCommand;
        global.Option = window.Option;
        document.execCommand = function () {
            const range = document.createRange();
            range.selectNodeContents(document.activeElement);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            return true;
        };

        try {
            const instance = jspreadsheet(root, {
                labels: { editCell: 'Zelle {0} bearbeiten' },
                worksheets: [{ data: [['a', '']], columns: [{}, { type: 'html' }] }],
            });
            const worksheet = instance[0];

            worksheet.openEditor(worksheet.records[0][1].element);

            const richText = worksheet.editor.cell.querySelector('[contenteditable]');
            expect(richText.getAttribute('role')).to.equal('textbox');
            expect(richText.getAttribute('aria-multiline')).to.equal('true');
            expect(richText.getAttribute('aria-label')).to.equal('Zelle B1 bearbeiten');
        } finally {
            global.Option = option;
            document.execCommand = execCommand;
        }
    });

    it('Translate the English names with setDictionary', () => {
        jspreadsheet.setDictionary({ 'Select all': 'Selecionar tudo', 'Edit cell {0}': 'Editar célula {0}' });

        try {
            const instance = jspreadsheet(root, {
                labels: { clipboardHelper: 'Área de transferência' },
                worksheets: [{ minDimensions: [3, 3] }],
            });
            const worksheet = instance[0];

            expect(worksheet.headerContainer.children[0].getAttribute('aria-label')).to.equal('Selecionar tudo');
            expect(worksheet.clipboard.textarea.getAttribute('aria-label')).to.equal('Área de transferência');

            worksheet.openEditor(worksheet.records[0][1].element);

            expect(worksheet.editor.cell.children[0].getAttribute('aria-label')).to.equal('Editar célula B1');
        } finally {
            delete document.dictionary['Select all'];
            delete document.dictionary['Edit cell {0}'];
        }
    });
});

describe('Keyboard focus', () => {
    const keydown = function (target, key, keyCode) {
        target.dispatchEvent(
            new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, composed: true, key: key, keyCode: keyCode, which: keyCode })
        );
    };

    const getTabStops = function (worksheet) {
        return Array.prototype.slice.call(worksheet.table.querySelectorAll('[tabindex="0"]'));
    };

    const wait = function () {
        return new Promise(function (resolve) {
            setTimeout(resolve, 10);
        });
    };

    it('Make the first cell the tab stop', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[0][0].element]);
        expect(worksheet.table.getAttribute('aria-multiselectable')).to.equal('true');
    });

    it('Mark the selected cells with aria-selected', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 3] }] })[0];
        const getCells = function (value) {
            const cells = worksheet.table.querySelectorAll('[aria-selected="' + value + '"]');
            return Array.prototype.map.call(cells, (td) => td.getAttribute('data-x') + ',' + td.getAttribute('data-y'));
        };

        // The tab stop is not selected
        expect(getCells('false')).to.eql(['0,0']);
        expect(getCells('true')).to.eql([]);

        worksheet.updateSelectionFromCoords(0, 0, 1, 1);

        expect(getCells('true')).to.eql(['0,0', '1,0', '0,1', '1,1']);
        expect(getCells('false')).to.eql([]);

        worksheet.updateSelectionFromCoords(2, 2, 2, 2);

        expect(getCells('true')).to.eql(['2,2']);
        expect(getCells('false')).to.eql([]);

        worksheet.resetSelection();

        expect(getCells('true')).to.eql([]);
        expect(getCells('false')).to.eql(['2,2']);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);

        expect(getCells('true')).to.eql(['1,1']);
        expect(getCells('false')).to.eql([]);
    });

    it('Move the tab stop and the focus', () => {
        const data = [
            ['0-0', true, '#ff0000', ''],
            ['1-0', false, '#00ff00', 'data:image/png;base64,iVBORw0KGgo='],
        ];
        const columns = [{}, { type: 'checkbox' }, { type: 'color', render: 'square' }, { type: 'image' }];
        const worksheet = jspreadsheet(root, { worksheets: [{ data: data, columns: columns }] })[0];
        const a1 = worksheet.records[0][0].element;
        const a2 = worksheet.records[1][0].element;

        // The focus on the spreadsheet container is in the grid
        worksheet.parent.element.focus();
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);

        expect(document.activeElement).to.equal(a1);
        expect(a1.getAttribute('aria-label')).to.equal('A1 0-0');

        keydown(a1, 'ArrowDown', 40);

        expect(document.activeElement).to.equal(a2);
        expect(getTabStops(worksheet)).to.deep.equal([a2]);
        expect(a1.tabIndex).to.equal(-1);
        expect(a1.hasAttribute('aria-label')).to.equal(false);
        expect(a2.getAttribute('aria-label')).to.equal('A2 1-0');

        worksheet.setValue('A2', 'Total');

        expect(a2.getAttribute('aria-label')).to.equal('A2 Total');

        // A cell without text gives its value
        keydown(a2, 'ArrowRight', 39);

        expect(document.activeElement.getAttribute('aria-label')).to.equal('B2 false');

        keydown(document.activeElement, 'ArrowRight', 39);

        expect(document.activeElement.getAttribute('aria-label')).to.equal('C2 #00ff00');

        // The value of an image is the image data
        keydown(document.activeElement, 'ArrowRight', 39);

        expect(document.activeElement.getAttribute('aria-label')).to.equal('D2');
    });

    it('Keep the focus outside the grid', () => {
        // The grid never had the focus
        document.activeElement.blur();
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];
        const input = document.createElement('input');
        root.appendChild(input);

        expect(document.activeElement).to.equal(document.body);

        worksheet.updateSelectionFromCoords(2, 2, 2, 2);

        expect(document.activeElement).to.equal(document.body);
        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[2][2].element]);

        input.focus();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);

        expect(document.activeElement).to.equal(input);
        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[1][1].element]);

        // The new cells do not take the focus from the input
        worksheet.records[1][1].element.focus();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        input.focus();
        worksheet.setData(createData(3, 3));

        expect(document.activeElement).to.equal(input);

        // A press outside the spreadsheet
        worksheet.records[1][1].element.focus();
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        document.activeElement.blur();
        document.body.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));
        worksheet.setData(createData(3, 3));

        expect(document.activeElement).to.equal(document.body);

        // The focus goes to an element outside the spreadsheet, and the element is removed
        const dialog = document.createElement('input');
        document.body.appendChild(dialog);
        worksheet.records[1][1].element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, cancelable: true }));

        expect(document.activeElement).to.equal(worksheet.records[1][1].element);

        dialog.focus();
        dialog.remove();
        worksheet.setData(createData(3, 3));

        expect(document.activeElement).to.equal(document.body);
    });

    it('Move the focus to the editor and back', async () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];
        const cell = worksheet.records[1][1].element;

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.openEditor(cell);
        // A change of the rows while the editor is open
        worksheet.insertRow();

        expect(cell.hasAttribute('tabindex')).to.equal(false);
        expect(cell.hasAttribute('aria-label')).to.equal(false);
        expect(getTabStops(worksheet)).to.deep.equal([]);

        worksheet.editor.cell.children[0].value = 'new';
        worksheet.closeEditor(cell, true);

        expect(cell.tabIndex).to.equal(0);
        expect(cell.getAttribute('aria-label')).to.equal('B2 new');

        cell.focus();
        keydown(cell, 'F2', 113);

        const editor = worksheet.editor.cell.children[0];
        expect(document.activeElement).to.equal(editor);

        keydown(editor, 'Escape', 27);
        await wait();

        expect(worksheet.editor.cell).to.equal(null);
        expect(document.activeElement).to.equal(cell);
        expect(cell.tabIndex).to.equal(0);
        expect(cell.getAttribute('aria-label')).to.equal('B2 new');

        // The editor of a cell that is not the tab stop
        worksheet.openEditor(worksheet.records[0][2].element);
        worksheet.closeEditor(worksheet.records[0][2].element, false);

        expect(getTabStops(worksheet)).to.deep.equal([cell]);
    });

    it('Focus the dropdown and custom editors', () => {
        const editorType = {
            createCell: function (cell, value) {
                cell.textContent = value;
                return cell;
            },
            updateCell: function (cell, value) {
                cell.textContent = value;
                return value;
            },
            openEditor: function (cell) {
                const control = document.createElement('input');
                cell.classList.add('editor');
                cell.innerHTML = '';
                cell.appendChild(control);
                control.focus();
            },
            closeEditor: function (cell) {
                return cell.children[0].value;
            },
        };
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b', 'c']], columns: [{}, { type: 'dropdown', source: ['a', 'b'] }, { type: editorType }] }],
        })[0];
        const dropdownCell = worksheet.records[0][1].element;
        const customCell = worksheet.records[0][2].element;

        dropdownCell.focus();
        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        keydown(dropdownCell, 'F2', 113);

        expect(document.activeElement).to.equal(dropdownCell.querySelector('.jdropdown-header'));
        expect(document.activeElement.getAttribute('aria-label')).to.equal('Edit cell B1');

        worksheet.closeEditor(dropdownCell, false);
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);
        worksheet.openEditor(customCell);

        expect(customCell.hasAttribute('tabindex')).to.equal(false);

        keydown(worksheet.editor.cell.children[0], 'Escape', 27);

        expect(worksheet.editor.cell).to.equal(null);
        expect(customCell.tabIndex).to.equal(0);
    });

    it('Focus a cell after a press', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];
        const cell = worksheet.records[2][1].element;

        document.body.focus();
        cell.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));

        expect(document.activeElement).to.equal(cell);
        expect(getTabStops(worksheet)).to.deep.equal([cell]);
    });

    it('Focus the cell after a focus loss', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3), search: true }] })[0];
        const a1 = worksheet.records[0][0].element;

        a1.focus();
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        a1.blur();
        keydown(document.body, 'ArrowDown', 40);

        expect(document.activeElement).to.equal(worksheet.records[1][0].element);

        // The focused row is not in the results
        worksheet.search('2-1');

        expect(document.activeElement).to.equal(worksheet.records[2][0].element);

        // The table attaches the rows again
        worksheet.resetSearch();

        expect(document.activeElement).to.equal(worksheet.records[2][0].element);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.setData(createData(3, 3));

        expect(document.activeElement).to.equal(worksheet.records[1][1].element);

        // The sort moves the focused row
        worksheet.orderBy(0, 1);

        expect(document.activeElement).to.equal(worksheet.records[1][1].element);
        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[1][1].element]);

        // The focused cell is no longer the tab stop
        worksheet.records[0][0].element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, cancelable: true }));
        worksheet.resetSelection();
        worksheet.hideColumn(0);

        expect(document.activeElement).to.equal(worksheet.records[0][1].element);
    });

    it('Send the keys to the grid that has the focus', async () => {
        const first = document.createElement('div');
        const second = document.createElement('div');
        root.appendChild(first);
        root.appendChild(second);
        const a = jspreadsheet(first, { worksheets: [{ data: createData(3, 3) }] })[0];
        const b = jspreadsheet(second, { worksheets: [{ data: createData(3, 3) }] })[0];
        // The spreadsheets add their document listeners after they build their worksheets
        await wait();

        a.records[0][0].element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, cancelable: true }));

        expect(a.selection.range).to.eql([0, 0, 0, 0]);

        // Tab moves the focus to the tab stop of the second grid
        b.records[0][0].element.focus();

        expect(a.selection.range).to.equal(null);
        expect(b.selection.range).to.eql([0, 0, 0, 0]);

        keydown(b.records[0][0].element, 'ArrowDown', 40);

        expect(a.selection.range).to.equal(null);
        expect(b.selection.range).to.eql([0, 1, 0, 1]);
        expect(document.activeElement).to.equal(b.records[1][0].element);
    });

    it('Keep the keys of a control outside the grid', async () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];
        // root is the element of the spreadsheet, so the input goes next to it
        const input = document.createElement('input');
        document.body.appendChild(input);
        await wait();

        worksheet.records[0][0].element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 1, which: 1 }));
        document.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, cancelable: true }));

        // Tab moves the focus out of the grid
        input.focus();
        keydown(input, 'x', 88);
        keydown(input, 'ArrowDown', 40);
        keydown(input, 'Delete', 46);

        // Tab on the input moves the focus on (issue #27)
        const tab = new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, composed: true, key: 'Tab', keyCode: 9, which: 9 });
        input.dispatchEvent(tab);

        expect(tab.defaultPrevented).to.equal(false);
        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('0-0');
        expect(document.activeElement).to.equal(input);

        // The focus that falls to the body leaves the keys to the grid
        input.blur();
        keydown(document.body, 'ArrowDown', 40);

        expect(worksheet.selection.range).to.eql([0, 1, 0, 1]);

        input.remove();
    });

    it('Move the focus in a shadow root', async () => {
        const host = document.createElement('div');
        document.body.appendChild(host);
        const shadowRoot = host.attachShadow({ mode: 'open' });
        const container = document.createElement('div');
        shadowRoot.appendChild(container);

        try {
            const worksheet = jspreadsheet(container, { root: shadowRoot, worksheets: [{ data: createData(3, 3) }] })[0];
            const a1 = worksheet.records[0][0].element;
            const a2 = worksheet.records[1][0].element;

            a1.focus();
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            keydown(a1, 'ArrowDown', 40);

            expect(shadowRoot.activeElement).to.equal(a2);

            keydown(a2, 'F2', 113);
            keydown(worksheet.editor.cell.children[0], 'Escape', 27);
            await wait();

            expect(shadowRoot.activeElement).to.equal(a2);
        } finally {
            host.remove();
        }
    });

    it('Skip hidden cells', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3), columns: [{ type: 'hidden' }, {}, {}] }] })[0];

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[0][1].element]);

        worksheet.hideRow(0);

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[1][1].element]);

        worksheet.hideColumn(1);

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[1][2].element]);
    });

    it('Move the tab stop out of removed cells', async () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(4, 3), pagination: 2, search: true }] })[0];
        const b2 = worksheet.records[1][1].element;

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.page(1);

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[2][0].element]);

        worksheet.page(0);

        expect(getTabStops(worksheet)).to.deep.equal([b2]);

        worksheet.resetSelection();
        worksheet.insertRow();

        expect(getTabStops(worksheet)).to.deep.equal([b2]);

        worksheet.deleteColumn(1);

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[0][0].element]);

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.resetSelection();
        worksheet.setData(createData(4, 2));

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[0][0].element]);

        // The selected row is not in the results
        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        worksheet.search('3-1');

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[3][0].element]);

        // Lazy loading removes the row of the tab stop
        const container = document.createElement('div');
        root.appendChild(container);
        const lazy = jspreadsheet(container, { worksheets: [{ data: createData(300, 2), lazyLoading: true, tableOverflow: true, tableHeight: '300px' }] })[0];
        lazy.content.dispatchEvent(new window.WheelEvent('wheel', { bubbles: true }));
        await new Promise(function (resolve) {
            setTimeout(resolve, 150);
        });

        expect(lazy.tbody.firstChild.getAttribute('data-y')).to.not.equal('0');
        expect(getTabStops(lazy)).to.deep.equal([lazy.tbody.firstChild.children[1]]);

        // After the sort, the first page does not have the row of the tab stop
        lazy.orderBy(0, 0);

        expect(getTabStops(lazy)).to.deep.equal([lazy.tbody.firstChild.children[1]]);
    });
});
