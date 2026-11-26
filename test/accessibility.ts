import { expect } from 'chai';

import type { MergeInfo } from '../src/types/internal';
import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

function createData(rows: number, columns: number) {
    const data: string[][] = [];
    for (let j = 0; j < rows; j++) {
        const row: string[] = [];
        for (let i = 0; i < columns; i++) {
            row.push(`${j}-${i}`);
        }
        data.push(row);
    }
    return data;
}

// Returns the aria-rowindex of the row of each data cell and the aria-colindex of the cell, as 'row:column'
function getIndexes(worksheet: Worksheet) {
    return worksheet.records.map((row, y) => {
        return row.map((record) => {
            return `${worksheet.rows[y].element.getAttribute('aria-rowindex')}:${record.element.getAttribute('aria-colindex')}`;
        });
    });
}

function getExpectedIndexes(rows: number, columns: number, headerRows?: number) {
    const result: string[][] = [];
    for (let j = 0; j < rows; j++) {
        const row: string[] = [];
        for (let i = 0; i < columns; i++) {
            row.push(`${j + 1 + (headerRows || 1)}:${i + 2}`);
        }
        result.push(row);
    }
    return result;
}

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

        const rows = Array.prototype.map.call(worksheet.tbody.children, (tr: Element) => tr.getAttribute('aria-rowindex'));
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
        // The cells under the merged cell, which the library keeps after the colspan and the rowspan
        const cells = (worksheet.options.mergeCells!.C1 as MergeInfo)[2]!;
        expect(cells.map((td) => `${td.getAttribute('data-x')}:${td.getAttribute('aria-colindex')}`)).to.eql(['3:5', '2:4', '3:5']);
    });

    it('Keep the row indexes after orderBy', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: [['c'], ['a'], ['b']] }],
        });
        const worksheet = instance[0];

        worksheet.orderBy(0, 0);

        expect(worksheet.getColumnData(0)).to.eql(['a', 'b', 'c']);
        expect(Array.prototype.map.call(worksheet.tbody.children, (tr: Element) => tr.getAttribute('aria-rowindex'))).to.eql(['2', '3', '4']);
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
        const footer = worksheet.table.querySelector('tfoot > tr')!;

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
        const footer = worksheet.table.querySelector('tfoot > tr')!;

        worksheet.hideColumn(1);

        // The cells of each row in the column D have the aria-colindex of the header of the column D
        const expected = worksheet.headers[3].getAttribute('aria-colindex');
        expect(expected).to.equal('5');
        expect(worksheet.filter!.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(footer.children[4].getAttribute('aria-colindex')).to.equal(expected);
        expect(worksheet.records[0][3].element.getAttribute('aria-colindex')).to.equal(expected);
        // The nested header G2 starts in the column C
        expect(nested.children[2].getAttribute('aria-colindex')).to.equal('4');
        expect(nested.children[0].getAttribute('aria-colindex')).to.equal('1');
        expect(worksheet.filter!.children[0].getAttribute('aria-colindex')).to.equal('1');
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

        expect(worksheet.editor.cell!.children[0].getAttribute('aria-label')).to.equal('Edit cell C2');
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

        expect(worksheet.editor.cell!.children[0].tagName).to.equal('TEXTAREA');
        expect(worksheet.editor.cell!.children[0].getAttribute('aria-label')).to.equal('Zelle A3 bearbeiten');

        worksheet.closeEditor(worksheet.editor.cell!, false);
        worksheet.openEditor(worksheet.records[0][1].element);

        expect(worksheet.editor.cell!.querySelector('.jdropdown-header')!.getAttribute('aria-label')).to.equal('Zelle B1 bearbeiten');

        worksheet.closeEditor(worksheet.editor.cell!, false);
        worksheet.openEditor(worksheet.records[2][3].element);

        expect(worksheet.editor.cell!.children[0].getAttribute('aria-label')).to.equal('Zelle D3 bearbeiten');
    });

    it('Give the rich text editor a role and a name', () => {
        // The html cells use the Option constructor of the window
        const option = global.Option;
        global.Option = window.Option;

        try {
            const instance = jspreadsheet(root, {
                labels: { editCell: 'Zelle {0} bearbeiten' },
                worksheets: [{ data: [['a', '']], columns: [{}, { type: 'html' }] }],
            });
            const worksheet = instance[0];

            worksheet.openEditor(worksheet.records[0][1].element);

            const richText = worksheet.editor.cell!.querySelector('[contenteditable]')!;
            expect(richText.getAttribute('role')).to.equal('textbox');
            expect(richText.getAttribute('aria-multiline')).to.equal('true');
            expect(richText.getAttribute('aria-label')).to.equal('Zelle B1 bearbeiten');
        } finally {
            global.Option = option;
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

            expect(worksheet.editor.cell!.children[0].getAttribute('aria-label')).to.equal('Editar célula B1');
        } finally {
            // An empty translation shows the English text again
            jspreadsheet.setDictionary({ 'Select all': '', 'Edit cell {0}': '' });
        }
    });

    it('Translate the names of the tabs with setDictionary (issue #32)', () => {
        jspreadsheet.setDictionary({ Worksheets: 'Arbeitsblätter', 'Tab Panels': 'Tabellenblätter', 'Add a new worksheet': 'Arbeitsblatt hinzufügen' });

        try {
            jspreadsheet(root, { tabs: true, worksheets: [{}, {}] });

            expect(root.querySelector('.jtabs-headers')!.getAttribute('aria-label')).to.equal('Arbeitsblätter');
            expect(root.querySelector('.jtabs-content')!.getAttribute('aria-label')).to.equal('Tabellenblätter');
            expect(root.querySelector('.jtabs-add')!.getAttribute('aria-label')).to.equal('Arbeitsblatt hinzufügen');
        } finally {
            // An empty translation shows the English text again
            jspreadsheet.setDictionary({ Worksheets: '', 'Tab Panels': '', 'Add a new worksheet': '' });
        }
    });

    it('Use the labels option for the names of the tabs', () => {
        jspreadsheet(root, {
            tabs: true,
            labels: { worksheets: 'Arbeitsblätter', tabPanels: 'Tabellenblätter', addWorksheet: 'Arbeitsblatt hinzufügen' },
            worksheets: [{}, {}],
        });

        expect(root.querySelector('.jtabs-headers')!.getAttribute('aria-label')).to.equal('Arbeitsblätter');
        expect(root.querySelector('.jtabs-content')!.getAttribute('aria-label')).to.equal('Tabellenblätter');
        expect(root.querySelector('.jtabs-add')!.getAttribute('aria-label')).to.equal('Arbeitsblatt hinzufügen');
    });
});

describe('Notes and announcements', () => {
    it('adds the note of the cellNote option to the name of the active cell', () => {
        const calls: [Jspreadsheet.WorksheetInstance, number, number][] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['12', 'a']] }],
            cellNote: (instance, x, y) => {
                calls.push([instance, x, y]);
                return x === 0 ? 'cached, not recalculated' : '';
            },
        });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        expect(worksheet.records[0][0].element.getAttribute('aria-label')).to.equal('A1 12, cached, not recalculated');
        expect(calls[calls.length - 1]).to.eql([worksheet, 0, 0]);

        // An empty note adds nothing
        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        expect(worksheet.records[0][1].element.getAttribute('aria-label')).to.equal('B1 a');
    });

    it('sets the name of the active cell again with labelCell after the note changes', () => {
        let note = '';
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['12', 'a']] }],
            cellNote: () => note,
        });
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        expect(worksheet.records[0][0].element.getAttribute('aria-label')).to.equal('A1 12');

        note = 'not recalculated';
        worksheet.labelCell(0, 0);
        expect(worksheet.records[0][0].element.getAttribute('aria-label')).to.equal('A1 12, not recalculated');

        // Only the active cell has a name
        worksheet.labelCell(1, 0);
        expect(worksheet.records[0][1].element.hasAttribute('aria-label')).to.equal(false);
    });

    it('reads a text of announce through a polite live region of the spreadsheet', async () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2] }] });
        await new Promise((resolve) => setTimeout(resolve));

        const regions = root.querySelectorAll('.jss_live');
        expect(regions.length).to.equal(1);
        expect(regions[0].getAttribute('role')).to.equal('status');
        expect(regions[0].getAttribute('aria-live')).to.equal('polite');
        expect(regions[0].textContent).to.equal('');

        worksheet.parent.announce('B4 changed by another user');
        expect(regions[0].textContent).to.equal('B4 changed by another user');

        // The same text again is a new element, so a screen reader reads it again
        const first = regions[0].firstElementChild;
        worksheet.parent.announce('B4 changed by another user');
        expect(regions[0].children.length).to.equal(1);
        expect(regions[0].firstElementChild).to.not.equal(first);
    });

    it('gives the note to the name of the first cell of a merged range, and again after the editor closes', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['12', 'b'],
                        ['c', 'd'],
                    ],
                    mergeCells: { A1: [2, 2] },
                },
            ],
            cellNote: (instance, x, y) => `note ${x},${y}`,
        });

        worksheet.updateSelectionFromCoords(1, 1, 1, 1);
        const cell = worksheet.records[0][0].element;
        expect(cell.getAttribute('aria-label')).to.equal('A1 12, note 0,0');

        worksheet.openEditor(cell);
        expect(cell.hasAttribute('aria-label')).to.equal(false);
        worksheet.closeEditor(cell, false);
        expect(cell.getAttribute('aria-label')).to.equal('A1 12, note 0,0');
    });
});

describe('Keyboard focus', () => {
    function keydown(target: Element, key: string, keyCode: number) {
        target.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, composed: true, key, keyCode, which: keyCode }));
    }

    function getTabStops(worksheet: Worksheet) {
        return Array.prototype.slice.call(worksheet.table.querySelectorAll('[tabindex="0"]')) as Element[];
    }

    function wait() {
        return new Promise((resolve) => {
            setTimeout(resolve, 10);
        });
    }

    it('Make the first cell the tab stop', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];

        expect(getTabStops(worksheet)).to.deep.equal([worksheet.records[0][0].element]);
        expect(worksheet.table.getAttribute('aria-multiselectable')).to.equal('true');
    });

    it('Mark the selected cells with aria-selected', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 3] }] })[0];
        function getCells(value: string) {
            const cells = worksheet.table.querySelectorAll(`[aria-selected="${value}"]`);
            return Array.prototype.map.call(cells, (td: Element) => `${td.getAttribute('data-x')},${td.getAttribute('data-y')}`);
        }

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
        const columns: Jspreadsheet.Column[] = [{}, { type: 'checkbox' }, { type: 'color', render: 'square' }, { type: 'image' }];
        const worksheet = jspreadsheet(root, { worksheets: [{ data, columns }] })[0];
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

        expect(document.activeElement!.getAttribute('aria-label')).to.equal('B2 false');

        keydown(document.activeElement!, 'ArrowRight', 39);

        expect(document.activeElement!.getAttribute('aria-label')).to.equal('C2 #00ff00');

        // The value of an image is the image data
        keydown(document.activeElement!, 'ArrowRight', 39);

        expect(document.activeElement!.getAttribute('aria-label')).to.equal('D2');
    });

    it('Keep the focus outside the grid', async () => {
        // The grid never had the focus
        (document.activeElement as HTMLElement).blur();
        const worksheet = jspreadsheet(root, { worksheets: [{ data: createData(3, 3) }] })[0];
        const input = document.createElement('input');
        root.appendChild(input);
        // The spreadsheet adds its document listeners after it builds its worksheets
        await wait();

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
        (document.activeElement as HTMLElement).blur();
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

        (worksheet.editor.cell!.children[0] as HTMLInputElement).value = 'new';
        worksheet.closeEditor(cell, true);

        expect(cell.tabIndex).to.equal(0);
        expect(cell.getAttribute('aria-label')).to.equal('B2 new');

        cell.focus();
        keydown(cell, 'F2', 113);

        const editor = worksheet.editor.cell!.children[0];
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
        const editorType: Jspreadsheet.CustomEditor = {
            createCell(cell, value) {
                cell.textContent = value as string;
                return cell;
            },
            updateCell(cell, value) {
                cell.textContent = value as string;
                return value;
            },
            openEditor(cell) {
                const control = document.createElement('input');
                cell.classList.add('editor');
                cell.innerHTML = '';
                cell.appendChild(control);
                control.focus();
            },
            closeEditor(cell) {
                return (cell.children[0] as HTMLInputElement).value;
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
        expect(document.activeElement!.getAttribute('aria-label')).to.equal('Edit cell B1');

        worksheet.closeEditor(dropdownCell, false);
        worksheet.updateSelectionFromCoords(2, 0, 2, 0);
        worksheet.openEditor(customCell);

        expect(customCell.hasAttribute('tabindex')).to.equal(false);

        keydown(worksheet.editor.cell!.children[0], 'Escape', 27);

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
            keydown(worksheet.editor.cell!.children[0], 'Escape', 27);
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
        lazy.content.dispatchEvent(new window.WheelEvent('wheel', { bubbles: true, deltaY: 100 }));
        await new Promise((resolve) => {
            setTimeout(resolve, 150);
        });

        expect((lazy.tbody.firstChild as HTMLElement).getAttribute('data-y')).to.not.equal('0');
        expect(getTabStops(lazy)).to.deep.equal([(lazy.tbody.firstChild as HTMLElement).children[1]]);

        // After the sort, the first page does not have the row of the tab stop
        lazy.orderBy(0, 0);

        expect(getTabStops(lazy)).to.deep.equal([(lazy.tbody.firstChild as HTMLElement).children[1]]);
    });
});

describe('Ids', () => {
    it('gives no two elements of a spreadsheet with two worksheets the same id (issue #68)', async () => {
        jspreadsheet(root, { tabs: true, worksheets: [{ minDimensions: [2, 2] }, { minDimensions: [2, 2] }] });
        await new Promise((resolve) => setTimeout(resolve));

        const ids = Array.from(root.querySelectorAll('[id]')).map((element) => element.id);

        expect(ids.filter((id, index) => ids.indexOf(id) !== index)).to.eql([]);
        expect(root.querySelectorAll('.jss_textarea').length).to.equal(2);
    });

    it('gives the tabs of two spreadsheets different ids', async () => {
        const other = document.createElement('div');
        document.body.appendChild(other);

        try {
            jspreadsheet(root, { tabs: true, worksheets: [{}, {}] });
            jspreadsheet(other, { tabs: true, worksheets: [{}, {}] });
            await new Promise((resolve) => setTimeout(resolve));

            const ids = Array.from(document.querySelectorAll('[role=tab], [role=tabpanel]')).map((element) => element.id);

            expect(ids.length).to.equal(8);
            expect(new Set(ids).size).to.equal(8);
            expect(ids.every((id) => id !== '')).to.equal(true);
        } finally {
            jspreadsheet.destroy(other as Jspreadsheet.JspreadsheetInstanceElement);
            other.remove();
        }
    });
});
