import { expect } from 'chai';
import type { Dropdown } from '../src/utils/dropdown';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

function rows(count: number) {
    const data: string[][] = [];
    for (let j = 0; j < count; j++) {
        data.push([`${j}`, '']);
    }
    return data;
}

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
        // The element of the editor holds its dropdown
        expect((cell.children[0] as HTMLElement & { dropdown: Dropdown }).dropdown.getValue(true)).to.deep.equal(['1']);

        worksheet.closeEditor(cell, true);
        expect(worksheet.editor.cell).to.equal(null);
        expect(`${worksheet.getValueFromCoords(0, 0)}`).to.equal('1');
    });

    it('saves and closes the text editor when lazy loading removes its row', async () => {
        // The grid loads rows 100 ms after a scroll, one load at a time. Let the loads of the tests before end.
        await new Promise((resolve) => setTimeout(resolve, 150));

        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];
        worksheet.openEditor(worksheet.records[0][1].element);
        worksheet.records[0][1].element.querySelector<HTMLInputElement | HTMLTextAreaElement>('input, textarea')!.value = 'typed';

        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
        await new Promise((resolve) => setTimeout(resolve, 150));

        expect(worksheet.tbody.contains(worksheet.records[0][1].element)).to.equal(false);
        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('typed');
    });
    describe('oneditordetach', () => {
        // The wheel loads rows 100 ms after the event, and the load removes the first rows, also the row of the editor
        async function detachByScroll(options: Partial<Jspreadsheet.SpreadsheetOptions>, ignoreEvents = false) {
            // The grid loads one page at a time. Let the loads of the tests before end.
            await new Promise((resolve) => setTimeout(resolve, 150));
            const worksheet = jspreadsheet(root, { worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }], ...options })[0];
            worksheet.openEditor(worksheet.records[0][1].element);
            worksheet.records[0][1].element.querySelector<HTMLInputElement | HTMLTextAreaElement>('input, textarea')!.value = 'typed';
            worksheet.parent.ignoreEvents = ignoreEvents;
            try {
                worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
                await new Promise((resolve) => setTimeout(resolve, 150));
            } finally {
                worksheet.parent.ignoreEvents = false;
            }
            return worksheet;
        }

        it('closes the editor without a save when the handler takes the edit', async () => {
            const calls: [number, number, string][] = [];
            const worksheet = await detachByScroll({
                oneditordetach: (instance, td, x, y) => {
                    calls.push([x, y, (td.children[0] as HTMLInputElement).value]);
                    return true;
                },
            });

            expect(calls).to.eql([[1, 0, 'typed']]);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(1, 0)).to.not.equal('typed');
        });

        it('saves the editor when the handler does not take the edit', async () => {
            let calls = 0;
            const worksheet = await detachByScroll({
                oneditordetach: () => {
                    calls++;
                    return undefined;
                },
            });

            expect(calls).to.equal(1);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(1, 0)).to.equal('typed');
        });

        it('does nothing more when the handler closed the editor itself', async () => {
            const ends: boolean[] = [];
            const worksheet = await detachByScroll({
                oneditordetach: (instance, td) => {
                    instance.closeEditor(td, true);
                    return undefined;
                },
                oneditionend: (instance, td, x, y, value, saved) => {
                    ends.push(saved);
                },
            });

            expect(ends).to.eql([true]);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(1, 0)).to.equal('typed');
        });

        it('does not save the edit when the handler moves the focus and takes the edit', async () => {
            const other = document.body.appendChild(document.createElement('input'));
            try {
                const worksheet = await detachByScroll({
                    oneditordetach: () => {
                        other.focus();
                        return true;
                    },
                });

                expect(document.activeElement).to.equal(other);
                expect(worksheet.getValueFromCoords(1, 0)).to.not.equal('typed');
            } finally {
                other.remove();
            }
        });

        it('gives the position of the cell after an insert above the editor', async () => {
            const calls: [number, number][] = [];
            await new Promise((resolve) => setTimeout(resolve, 150));
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
                oneditordetach: (instance, td, x, y) => {
                    calls.push([x, y]);
                    return true;
                },
            })[0];
            worksheet.openEditor(worksheet.records[2][1].element);

            worksheet.insertRow(3, 0, true);
            worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
            await new Promise((resolve) => setTimeout(resolve, 150));

            expect(calls).to.eql([[1, 5]]);
        });

        it('occurs when a move of the selection loads another page', () => {
            const calls: [number, number][] = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
                oneditordetach: (instance, td, x, y) => {
                    calls.push([x, y]);
                    return true;
                },
            })[0];
            worksheet.updateSelectionFromCoords(1, 0, 1, 0);
            worksheet.openEditor(worksheet.records[0][1].element);

            worksheet.last(false, true);

            expect(calls).to.eql([[1, 0]]);
            expect(worksheet.editor.isOpen()).to.equal(false);
        });

        it('occurs while ignoreEvents is set', async () => {
            let calls = 0;
            const worksheet = await detachByScroll(
                {
                    oneditordetach: () => {
                        calls++;
                        return true;
                    },
                },
                true
            );

            expect(calls).to.equal(1);
            expect(worksheet.getValueFromCoords(1, 0)).to.not.equal('typed');
        });
    });

    describe('closeEditor without an open editor', () => {
        it('does nothing for a cell that is not edited', () => {
            const events: string[] = [];
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

            a1.querySelector<HTMLInputElement | HTMLTextAreaElement>('input, textarea')!.value = 'typed';
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
            const events: string[] = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [['a']],
                        columns: [
                            {
                                type: {
                                    createCell: (td, value) => {
                                        // The values of the column are text
                                        td.textContent = value as string;
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
                                        // The values of the column are text
                                        td.textContent = value as string;
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
    const create = (options?: Jspreadsheet.WorksheetOptions) =>
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
        const calls: string[][] = [];
        const worksheet = create({
            columns: [
                {
                    type: 'dropdown',
                    source: ['a', 'c'],
                    filter: (element, cell, x, y, source) => {
                        calls.push([x, y]);
                        return source!;
                    },
                },
            ],
        });

        worksheet.openEditor(worksheet.getCell('A2'));

        expect(calls).to.deep.equal([['0', '1']]);
        expect(worksheet.editor.x).to.equal(0);
        expect(worksheet.editor.y).to.equal(1);
    });

    it('saves the value typed in the text editor', () => {
        const worksheet = create();
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);
        (cell.children[0] as HTMLInputElement).value = 'x';
        worksheet.closeEditor(cell, true);

        expect(worksheet.getValue('A1')).to.equal('x');
        expect(cell.classList.contains('editor')).to.equal(false);
        expect(worksheet.history.records.length).to.equal(1);
    });

    it('opens the text editor without the value when empty is true', () => {
        const worksheet = create();
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell, true);

        expect((cell.children[0] as HTMLInputElement).value).to.equal('');
    });

    it('saves 0 for an empty numeric cell, or an empty value with allowEmpty', () => {
        const worksheet = create({ columns: [{ type: 'numeric' }, { type: 'numeric', allowEmpty: true }], data: [[1, 2]] });

        worksheet.openEditor(worksheet.getCell('A1'));
        (worksheet.getCell('A1').children[0] as HTMLInputElement).value = '';
        worksheet.closeEditor(worksheet.getCell('A1'), true);

        worksheet.openEditor(worksheet.getCell('B1'));
        (worksheet.getCell('B1').children[0] as HTMLInputElement).value = '';
        worksheet.closeEditor(worksheet.getCell('B1'), true);

        expect(worksheet.options.data[0]).to.deep.equal([0, '']);
    });

    it('toggles a checkbox when its editor opens', () => {
        const worksheet = create({ columns: [{ type: 'checkbox' }], data: [[false]] });

        worksheet.openEditor(worksheet.getCell('A1'));

        expect(worksheet.options.data[0][0]).to.equal(true);
    });

    it('dispatches the events of the edition', () => {
        const events: (Jspreadsheet.CellValue | undefined)[][] = [];
        // The events are options of the spreadsheet
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            oneditionstart: (instance, cell, x, y) => events.push(['start', x, y]),
            oncreateeditor: (instance, cell, x, y) => events.push(['create', x, y]),
            oneditionend: (instance, cell, x, y, value, save) => events.push(['end', x, y, value, save]),
        })[0];
        const cell = worksheet.getCell('B1');

        worksheet.openEditor(cell);
        (cell.children[0] as HTMLInputElement).value = 'z';
        worksheet.closeEditor(cell, true);

        expect(events).to.deep.equal([
            ['start', 1, 0],
            ['create', 1, 0],
            ['end', 1, 0, 'z', true],
        ]);
    });
});

describe('The text editor and line breaks', () => {
    function altEnter(editor: HTMLElement) {
        editor.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter', keyCode: 13, which: 13, altKey: true }));
    }

    it('keeps the line breaks of a value when the editor saves it as it is', () => {
        const changes: unknown[] = [];
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: [['a\nb']] }],
            onchange: (instance, cell, x, y, value) => {
                changes.push(value);
            },
        });
        const cell = worksheet.getCell('A1');

        worksheet.openEditor(cell);
        expect((cell.children[0] as HTMLTextAreaElement).value).to.equal('a\nb');
        worksheet.closeEditor(cell, true);

        expect(worksheet.getValue('A1')).to.equal('a\nb');
        expect(changes).to.eql([]);
    });

    it('shows one line without wrap in a cell without word wrap, and keeps the wrap of word wrap', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: [['a', 'b']], columns: [{}, { wordWrap: true }] }] });

        worksheet.openEditor(worksheet.getCell('A1'));
        const editor = worksheet.getCell('A1').children[0];
        expect([editor.tagName, editor.getAttribute('rows'), editor.getAttribute('wrap')]).to.eql(['TEXTAREA', '1', 'off']);
        worksheet.closeEditor(worksheet.getCell('A1'), false);

        worksheet.openEditor(worksheet.getCell('B1'));
        const wrapped = worksheet.getCell('B1').children[0];
        expect([wrapped.tagName, wrapped.getAttribute('rows'), wrapped.getAttribute('wrap')]).to.eql(['TEXTAREA', null, null]);
    });

    it('adds a line break at the caret with Alt+Enter in a cell without word wrap, and replaces the selected text', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: [['abcd']] }] });
        const cell = worksheet.getCell('A1');
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.openEditor(cell);
        const editor = cell.children[0] as HTMLTextAreaElement;
        const inputs: string[] = [];
        editor.addEventListener('input', () => inputs.push(editor.value));

        editor.setSelectionRange(1, 3);
        altEnter(editor);

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(editor.value).to.equal('a\nd');
        expect([editor.selectionStart, editor.selectionEnd]).to.eql([2, 2]);
        expect(inputs).to.eql(['a\nd']);

        worksheet.closeEditor(cell, true);
        expect(worksheet.getValue('A1')).to.equal('a\nd');
    });

    it('closes the editor of another type with Alt+Enter, as with Enter', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: [['#ff0000']], columns: [{ type: 'color' }] }] });
        const cell = worksheet.getCell('A1');
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.openEditor(cell);
        const editor = cell.children[0] as HTMLInputElement;
        let blurred = false;
        editor.addEventListener('blur', () => {
            blurred = true;
        });
        editor.focus();

        altEnter(editor);

        expect(blurred).to.equal(true);
        // The color picker closes on its own events, which jsdom does not send
        worksheet.closeEditor(cell, false);
    });
});
