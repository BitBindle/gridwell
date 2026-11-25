const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const rows = function (count) {
    const data = [];
    for (let j = 0; j < count; j++) {
        data.push(['' + j, '']);
    }
    return data;
};

describe('Editor', () => {
    it('keeps the text editor open when the grid scrolls', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[1][0].element);
        worksheet.content.dispatchEvent(new window.Event('scroll'));

        expect(worksheet.editor.cell).to.not.equal(null);
    });

    it('closes a dropdown editor when the grid scrolls', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: rows(300),
                    columns: [{ type: 'dropdown', source: ['0', '1', '2'] }, {}],
                    lazyLoading: true,
                    tableOverflow: true,
                },
            ],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[1][0].element);
        worksheet.content.dispatchEvent(new window.Event('scroll'));

        expect(worksheet.editor.cell).to.equal(null);
    });

    it('opens a multiple dropdown editor when the value is a number', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [[1], ['1;2']],
                    columns: [
                        {
                            type: 'dropdown',
                            multiple: true,
                            source: [
                                { id: 1, name: 'One' },
                                { id: 2, name: 'Two' },
                            ],
                        },
                    ],
                },
            ],
        });

        const worksheet = instance[0];
        const cell = worksheet.getCell('A1');
        worksheet.openEditor(cell);

        expect(worksheet.editor.cell).to.not.equal(null);
        expect(cell.children[0].dropdown.getValue(true)).to.deep.equal(['1']);

        worksheet.closeEditor(cell, true);
        expect(worksheet.editor.cell).to.equal(null);
        expect('' + worksheet.getValueFromCoords(0, 0)).to.equal('1');
    });

    it('saves and closes the text editor when lazy loading removes its row', async () => {
        // The grid loads rows 100 ms after a scroll, one load at a time. Let the loads of the tests before end.
        await new Promise((resolve) => setTimeout(resolve, 150));

        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[0][1].element);
        worksheet.records[0][1].element.querySelector('input, textarea').value = 'typed';

        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
        await new Promise((resolve) => setTimeout(resolve, 150));

        expect(worksheet.tbody.contains(worksheet.records[0][1].element)).to.equal(false);
        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('typed');
    });
    describe('closeEditor without an open editor', () => {
        it('does nothing for a cell that is not edited', () => {
            const events = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['a', 'b']] }],
                oneditionend: () => events.push('oneditionend'),
            })[0];
            const cell = worksheet.getCell('A1');

            worksheet.closeEditor(cell, false);
            worksheet.closeEditor(cell, true);

            // No editor was opened
            expect(worksheet.editor.cell).to.equal(null);
            expect(cell.textContent).to.equal('a');
            expect(worksheet.getData()).to.deep.equal([['a', 'b']]);
            expect(events).to.deep.equal([]);
        });

        it('keeps the editor of another cell open', () => {
            const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a', 'b']] }] })[0];
            const a1 = worksheet.getCell('A1');
            const b1 = worksheet.getCell('B1');

            worksheet.openEditor(a1);
            worksheet.closeEditor(b1, true);

            expect(worksheet.editor.cell).to.equal(a1);
            expect(a1.classList.contains('editor')).to.equal(true);
            expect(b1.textContent).to.equal('b');

            a1.querySelector('input, textarea').value = 'typed';
            worksheet.closeEditor(a1, true);
            expect(worksheet.getData()).to.deep.equal([['typed', 'b']]);
        });
    });

    describe('Cells without an editor', () => {
        it('does not edit a cell of type hidden', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['id-7', '']], columns: [{ type: 'hidden' }, { type: 'text' }] }],
            })[0];
            const cell = worksheet.getCell('A1');

            worksheet.openEditor(cell);
            expect(worksheet.editor.cell).to.equal(null);

            worksheet.closeEditor(cell, false);
            worksheet.closeEditor(cell, true);
            expect(worksheet.editor.cell).to.equal(null);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('id-7');
        });

        it('does not edit a cell of a custom type without openEditor', () => {
            const events = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a']],
                        columns: [
                            {
                                type: {
                                    createCell: (td, value) => {
                                        td.textContent = value;
                                        return td;
                                    },
                                },
                            },
                        ],
                    },
                ],
                oncreateeditor: () => events.push('oncreateeditor'),
            })[0];
            const cell = worksheet.getCell('A1');

            worksheet.openEditor(cell);

            expect(worksheet.editor.cell).to.equal(null);
            expect(cell.classList.contains('editor')).to.equal(false);
            expect(events).to.deep.equal([]);
        });

        it('closes the editor of a custom type without closeEditor and keeps the value', () => {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a'], ['b']],
                        columns: [
                            {
                                type: {
                                    createCell: (td, value) => {
                                        td.textContent = value;
                                        return td;
                                    },
                                    openEditor: (td) => {
                                        td.classList.add('editor');
                                        td.innerHTML = '<input>';
                                    },
                                },
                            },
                        ],
                    },
                ],
            })[0];

            const a1 = worksheet.getCell('A1');
            worksheet.openEditor(a1);
            expect(worksheet.editor.cell).to.equal(a1);
            worksheet.closeEditor(a1, true);
            expect(worksheet.editor.cell).to.equal(null);
            expect(a1.textContent).to.equal('a');

            const a2 = worksheet.getCell('A2');
            worksheet.openEditor(a2);
            worksheet.closeEditor(a2, false);
            expect(worksheet.editor.cell).to.equal(null);
            expect(a2.textContent).to.equal('b');
            expect(worksheet.getData()).to.deep.equal([['a'], ['b']]);
        });
    });
});

describe('The editor of a worksheet', () => {
    const create = (options) =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                    ...options,
                },
            ],
        })[0];

    it('is closed before an editor opens', () => {
        const worksheet = create();

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.editor.x).to.equal(-1);
        expect(worksheet.editor.y).to.equal(-1);
    });

    it('keeps the cell, its position and its HTML while the editor is open', () => {
        const worksheet = create();
        const cell = worksheet.getCell('B2');

        worksheet.openEditor(cell);

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(worksheet.editor.cell).to.equal(cell);
        expect(worksheet.editor.x).to.equal(1);
        expect(worksheet.editor.y).to.equal(1);
        expect(worksheet.editor.html).to.equal('d');

        worksheet.closeEditor(cell, false);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.editor.html).to.equal('');
        expect(cell.innerHTML).to.equal('d');
    });

    it('restores the HTML of the cell when the value does not change', () => {
        const worksheet = create();
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);
        worksheet.closeEditor(cell, true);

        expect(cell.innerHTML).to.equal('a');
        expect(worksheet.history.records.length).to.equal(0);
    });

    it('stays closed for a column type without an editor', () => {
        const worksheet = create({ columns: [{ type: 'checkbox' }, { type: 'hidden' }] });

        worksheet.openEditor(worksheet.getCell('A1'));
        expect(worksheet.editor.isOpen()).to.equal(false);

        worksheet.openEditor(worksheet.getCell('B1'));
        expect(worksheet.editor.isOpen()).to.equal(false);
    });

    it('gives the filter of a dropdown the column and the row as text', () => {
        const calls = [];
        const worksheet = create({
            columns: [
                {
                    type: 'dropdown',
                    source: ['a', 'c'],
                    filter: (element, cell, x, y, source) => {
                        calls.push([x, y]);
                        return source;
                    },
                },
            ],
        });

        worksheet.openEditor(worksheet.getCell('A2'));

        expect(calls).to.deep.equal([['0', '1']]);
        expect(worksheet.editor.x).to.equal(0);
        expect(worksheet.editor.y).to.equal(1);
    });
});
