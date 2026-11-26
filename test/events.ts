import { expect } from 'chai';

import type { WorksheetHistoryRecord } from '../src/types/internal';
import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

const load = (options: Jspreadsheet.SpreadsheetOptions, element?: HTMLElement) => {
    return new Promise<Worksheet>((resolve) => {
        const instance = jspreadsheet(element || root, {
            ...options,
            onload: () => resolve(instance[0]),
        });
    });
};

const press = (element: Element) => {
    element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
};

describe('Mouse events outside the grid', () => {
    it('keeps the selection when the press is on a child element of a jss_object control', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        const control = document.createElement('div');
        control.className = 'jss_object';
        control.innerHTML = '<span>Bold</span>';
        document.body.appendChild(control);

        press(worksheet.records[1][1].element);
        press(control.querySelector('span')!);
        control.remove();

        expect(worksheet.selection.range!.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('takes a press on a cell of a host table with the class jss_object without an error', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        const host = document.createElement('div');
        host.innerHTML = '<table><tbody><tr><td class="jss_object">Restore</td></tr></tbody></table>';
        document.body.appendChild(host);

        const errors: string[] = [];
        const listener = (event: ErrorEvent) => errors.push(String(event.error || event.message));
        window.addEventListener('error', listener);

        press(worksheet.records[1][1].element);
        press(host.querySelector('td')!);

        window.removeEventListener('error', listener);
        host.remove();

        expect(errors).to.eql([]);
        expect(worksheet.selection.range!.map(Number)).to.eql([1, 1, 1, 1]);
    });
});

describe('Keys on a jss_object control outside the grid', () => {
    const keydown = (target: EventTarget, key: string, keyCode: number) => {
        const event = new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, composed: true, key, keyCode, which: keyCode });
        target.dispatchEvent(event);
        return event;
    };

    const paste = (target: EventTarget, text: string) => {
        const event = new window.Event('paste', { bubbles: true, cancelable: true });
        // jsdom has no ClipboardEvent, and the handler reads only getData
        (event as Event & { clipboardData: { getData: () => string } }).clipboardData = { getData: () => text };
        target.dispatchEvent(event);
        return event;
    };

    it('keeps the selection and the default action of the keys on a jss_object button', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] });

        const button = document.createElement('input');
        button.type = 'button';
        button.className = 'jss_object';
        document.body.appendChild(button);

        press(worksheet.records[1][1].element);
        press(button);
        button.focus();

        const results = [keydown(button, 'ArrowDown', 40), keydown(button, 'Enter', 13), keydown(button, 'Tab', 9)].map((event) => event.defaultPrevented);
        button.remove();

        expect(results).to.eql([false, false, false]);
        expect(worksheet.selection.range!.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('keeps the selection with the keys on a jss_object button in a shadow root', async () => {
        const host = document.createElement('div');
        document.body.appendChild(host);
        const shadowRoot = host.attachShadow({ mode: 'open' });
        const container = document.createElement('div');
        const button = document.createElement('input');
        button.type = 'button';
        button.className = 'jss_object';
        shadowRoot.appendChild(container);
        shadowRoot.appendChild(button);

        try {
            const worksheet = await load({ root: shadowRoot, worksheets: [{ minDimensions: [3, 4] }] }, container);

            press(worksheet.records[1][1].element);
            button.focus();

            const event = keydown(button, 'ArrowDown', 40);

            expect(event.defaultPrevented).to.equal(false);
            expect(worksheet.selection.range!.map(Number)).to.eql([1, 1, 1, 1]);
        } finally {
            host.remove();
        }
    });

    it('keeps the cells with a paste on a jss_object text input', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']] }] });

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'jss_object';
        document.body.appendChild(input);

        press(worksheet.records[0][0].element);
        press(input);
        input.focus();

        const event = paste(input, 'c');
        input.remove();

        expect(event.defaultPrevented).to.equal(false);
        expect(worksheet.getValue('A1')).to.equal('a');
    });

    it('takes the keys and a paste on a worksheet in a jss_object element', async () => {
        const wrapper = document.createElement('div');
        wrapper.className = 'jss_object';
        const container = document.createElement('div');
        wrapper.appendChild(container);
        document.body.appendChild(wrapper);

        try {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] }, container);

            press(worksheet.records[1][1].element);
            container.focus();

            const results = [keydown(container, 'ArrowDown', 40), keydown(container, 'ArrowRight', 39)].map((event) => event.defaultPrevented);
            expect(results).to.eql([true, true]);
            expect(worksheet.selection.range!.map(Number)).to.eql([2, 2, 2, 2]);

            paste(container, 'c');
            expect(worksheet.getValue('C3')).to.equal('c');
        } finally {
            wrapper.remove();
        }
    });

    it('takes the keys and a paste on an element without the class jss_object', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] });

        press(worksheet.records[1][1].element);

        const event = keydown(document.body, 'ArrowDown', 40);
        expect(event.defaultPrevented).to.equal(true);
        expect(worksheet.selection.range!.map(Number)).to.eql([1, 2, 1, 2]);

        paste(document.body, 'c');
        expect(worksheet.getValue('B3')).to.equal('c');
    });

    it('closes a custom editor with a jss_object control with Escape', async () => {
        // A custom editor with a jss_object control outside the worksheet, for example a floating panel
        const control = document.createElement('input');
        control.className = 'jss_object';
        const editorType: Jspreadsheet.CustomEditor = {
            createCell(cell, value) {
                // The cells of the test have texts
                cell.textContent = value as string;
                return cell;
            },
            updateCell(cell, value) {
                cell.textContent = value as string;
                return value;
            },
            openEditor(cell) {
                cell.classList.add('editor');
                document.body.appendChild(control);
                control.focus();
            },
            closeEditor() {
                control.remove();
                return control.value;
            },
        };
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']], columns: [{}, { type: editorType }] }] });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        worksheet.openEditor(worksheet.records[0][1].element);
        control.value = 'c';

        keydown(control, 'Escape', 27);
        control.remove();

        expect(worksheet.editor.cell).to.equal(null);
        expect(worksheet.getValue('B1')).to.equal('b');
    });
});

describe('Dispatch events', () => {
    it('gives onevent and plugins the event name and the arguments, and the event handler the arguments', () => {
        const calls: unknown[][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            // onevent takes its arguments as a rest of a union of tuples, which a handler with named parameters does not fit (issue #133)
            onevent(...[event, instance, cell, x, y, value]) {
                if (event === 'onchange') {
                    calls.push(['onevent', instance === worksheet, x, y, value]);
                }
            },
            onchange(instance, cell, x, y, value) {
                calls.push(['onchange', instance === worksheet, x, y, value]);
            },
            plugins: {
                test() {
                    return {
                        // As onevent above (issue #133)
                        onevent(...[event, instance, cell, x, y, value]) {
                            if (event === 'onchange') {
                                calls.push(['plugin', instance === worksheet, x, y, value]);
                            }
                        },
                    };
                },
            },
        })[0];

        worksheet.setValueFromCoords(1, 0, 'c');

        expect(calls).to.eql([
            ['onevent', true, 1, 0, 'c'],
            ['onchange', true, 1, 0, 'c'],
            ['plugin', true, 1, 0, 'c'],
        ]);
    });

    it('gives onbeforechange and onchange undefined as the new value when setValue has no value', () => {
        const calls: unknown[][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            onbeforechange(instance, cell, x, y, value) {
                calls.push(['onbeforechange', value]);
            },
            onchange(instance, cell, x, y, value, oldValue) {
                calls.push(['onchange', value, oldValue]);
            },
        })[0];

        worksheet.setValue('A1');

        expect(calls).to.eql([
            ['onbeforechange', undefined],
            ['onchange', undefined, 'a'],
        ]);
    });

    it('returns the result of the handler, so that onbeforeinsertrow can cancel the insertion', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            onbeforeinsertrow() {
                return false;
            },
        })[0];

        worksheet.insertRow();

        expect(worksheet.getData()).to.eql([['a']]);
    });

    it('keeps a false of the handler when a plugin with onevent returns nothing or another value (issue #37)', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            onbeforeinsertrow: () => false,
            plugins: {
                logger: () => ({ onevent: () => {} }),
                // onevent takes its arguments as a rest of a union of tuples, which a handler with named parameters does not fit (issue #133)
                other: () => ({ onevent: (...[event]) => (event === 'onbeforeinsertrow' ? true : undefined) }),
            },
        })[0];

        worksheet.insertRow();

        expect(worksheet.getData()).to.eql([['a']]);
    });

    it('lets a plugin cancel an action, and keeps the result of the handler when onevent returns nothing (issue #37)', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            onevent: () => {},
            onbeforechange: (instance, cell, x, y, value) => `${value}!`,
            // onevent takes its arguments as a rest of a union of tuples, which a handler with named parameters does not fit (issue #133)
            plugins: { guard: () => ({ onevent: (...[event]) => (event === 'onbeforedeleterow' ? false : undefined) }) },
        })[0];

        worksheet.setValue('A1', 'b');
        worksheet.insertRow();
        worksheet.deleteRow(0);

        expect(worksheet.getData()).to.eql([['b!'], ['']]);
    });

    it('gives oncreateworksheet the options of the new worksheet', async () => {
        let createdOptions: Jspreadsheet.WorksheetOptions | null = null;
        const worksheet = await load({
            worksheets: [{ minDimensions: [2, 2], worksheetName: 'First' }],
        });

        const created = new Promise((resolve) => {
            worksheet.parent.config.oncreateworksheet = function (instance, options, index) {
                createdOptions = options;
                resolve(index);
            };
        });
        worksheet.createWorksheet({ minDimensions: [3, 3], worksheetName: 'Second' });

        expect(await created).to.equal(1);
        expect(createdOptions!.worksheetName).to.equal('Second');
        expect(createdOptions).to.equal(worksheet.parent.worksheets[1].options);
    });

    it('gives onresizerow the old height as a number after the row was resized with the mouse', () => {
        const calls: number[][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
            onresizerow(instance, row, newHeight, oldHeight) {
                calls.push([row, newHeight, oldHeight]);
            },
        })[0];

        // A resize with the mouse leaves the height attribute on the row
        worksheet.rows[0].element.setAttribute('height', 40);
        worksheet.setHeight(0, 50);

        expect(calls).to.eql([[0, 50, 40]]);
    });

    it('gives the change events the column and row as numbers, however the cell changed (issue #52)', () => {
        const calls: unknown[][] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 3] }],
            onbeforechange: (instance, cell, x, y) => calls.push(['before', x, y]),
            onchange: (instance, cell, x, y) => calls.push(['change', x, y]),
            onafterchanges: (instance, records) => calls.push(['after', records[0].x, records[0].y]),
        })[0];
        const expected = [
            ['before', 2, 1],
            ['change', 2, 1],
            ['after', 2, 1],
        ];

        // A typed edit
        const td = worksheet.records[1][2].element;
        worksheet.openEditor(td);
        (td.children[0] as HTMLInputElement).value = 'v';
        worksheet.closeEditor(td, true);
        expect(calls).to.eql(expected);
        expect((worksheet.history.records[0] as Extract<WorksheetHistoryRecord, { action: 'setValue' }>).records[0]).to.include({ x: 2, y: 1 });

        // A cell element and a list of cell elements
        calls.length = 0;
        worksheet.setValue(td, 'w');
        expect(calls).to.eql(expected);

        calls.length = 0;
        worksheet.setValue([td], 'x');
        expect(calls).to.eql(expected);

        // A cell name
        calls.length = 0;
        worksheet.setValue('C2', 'y');
        expect(calls).to.eql(expected);
    });

    it('cancels the change when onbeforechange returns false (issue #38)', () => {
        const calls: string[] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            onbeforechange: (instance, cell, x, y, value) => (value === 'no' ? false : undefined),
            onchange: () => calls.push('onchange'),
            onafterchanges: () => calls.push('onafterchanges'),
        })[0];

        worksheet.setValueFromCoords(0, 0, 'no');

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('a');
        expect(worksheet.getCell('A1').textContent).to.equal('a');
        expect(calls).to.eql([]);
        expect(worksheet.history.records.length).to.equal(0);

        // A typed edit that is cancelled leaves the text of the cell
        const td = worksheet.getCell('B1');
        worksheet.openEditor(td);
        (td.children[0] as HTMLInputElement).value = 'no';
        worksheet.closeEditor(td, true);

        expect(worksheet.getValueFromCoords(1, 0)).to.equal('b');
        expect(td.textContent).to.equal('b');
        expect(td.children.length).to.equal(0);
        expect(calls).to.eql([]);

        // Another value is written
        worksheet.setValueFromCoords(0, 0, 'yes');

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('yes');
    });

    it('gives no change, history entry or onafterchanges for a read-only cell (issue #58)', () => {
        const calls: (string | Jspreadsheet.CellChange[])[] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']], columns: [{ readOnly: true }, {}] }],
            onchange: () => calls.push('onchange'),
            onafterchanges: (instance, changes) => calls.push(changes),
        })[0];

        worksheet.setValue('A1', 'x');
        worksheet.setValueFromCoords(0, 0, 'y');

        expect(calls).to.eql([]);
        expect(worksheet.history.records.length).to.equal(0);
        expect(worksheet.getData()).to.eql([['a', 'b']]);

        // Only the cell that changed is in the records
        worksheet.setValue(['A1', 'B1'], 'z');

        expect(calls.length).to.equal(2);
        expect(calls[1]).to.eql([{ x: 1, y: 0, value: 'z', oldValue: 'b' }]);
        expect(worksheet.history.records.length).to.equal(1);
        expect((worksheet.history.records[0] as Extract<WorksheetHistoryRecord, { action: 'setValue' }>).records.length).to.equal(1);
    });
});

describe('Destroy', () => {
    it('removes every event handler that the spreadsheet added to its root', async () => {
        const container = document.createElement('div');
        document.body.appendChild(container);

        const added: [string, EventListenerOrEventListenerObject][] = [];
        const removed: [string, EventListenerOrEventListenerObject][] = [];
        const addEventListener = container.addEventListener;
        const removeEventListener = container.removeEventListener;
        container.addEventListener = function (type: string, handler: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions) {
            added.push([type, handler]);
            return addEventListener.call(this, type, handler, options);
        };
        container.removeEventListener = function (type: string, handler: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions) {
            removed.push([type, handler]);
            return removeEventListener.call(this, type, handler, options);
        };

        await load({ root: container, worksheets: [{ minDimensions: [3, 3] }] });
        removed.length = 0;

        // root is the element of the spreadsheet
        jspreadsheet.destroy(root as Jspreadsheet.JspreadsheetInstanceElement, true);
        container.remove();

        expect(added.length).to.be.above(0);
        for (const [type, handler] of added) {
            expect(
                removed.some(([t, h]) => t === type && h === handler),
                type
            ).to.equal(true);
        }
    });
});

describe('Context menu', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    const contextmenu = (element: Element) => element.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

    it('selects the cell outside the selection and passes its position as text', async () => {
        const calls: unknown[][] = [];
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            contextMenu: (instance, x, y, e, items, role) => {
                calls.push([x, y, role]);
                return false;
            },
        });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        contextmenu(worksheet.records[2][1].element);

        expect(worksheet.selection.range).to.eql([1, 2, 1, 2]);
        expect(calls).to.eql([['1', '2', 'cell']]);
    });

    it('keeps a selection that has the cell', async () => {
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            contextMenu: () => false,
        });

        worksheet.updateSelectionFromCoords(0, 0, 2, 2);
        contextmenu(worksheet.records[1][1].element);

        expect(worksheet.selection.range).to.eql([0, 0, 2, 2]);
    });

    it('selects the row of a row number', async () => {
        const calls: unknown[][] = [];
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            contextMenu: (instance, x, y, e, items, role) => {
                calls.push([x, y, role]);
                return false;
            },
        });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        contextmenu(worksheet.rows[2].element.children[0]);

        expect(worksheet.selection.range).to.eql([0, 2, 2, 2]);
        expect(calls).to.eql([[null, '2', 'row']]);
    });

    it('selects the column of a header', async () => {
        const worksheet = await load({
            worksheets: [{ minDimensions: [3, 3] }],
            contextMenu: () => false,
        });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        contextmenu(worksheet.headers[2]);

        expect(worksheet.selection.range).to.eql([2, 0, 2, 2]);
    });
});

describe('Rename a column (issue #42)', () => {
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    const mouse = (element: EventTarget, type: string) => element.dispatchEvent(new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1 }));

    let prompt: typeof window.prompt;
    let answers: (string | null)[];
    let questions: [string | undefined, string | undefined][];

    beforeEach(() => {
        prompt = window.prompt;
        answers = [];
        questions = [];
        window.prompt = global.prompt = (text?: string, value?: string) => {
            questions.push([text, value]);
            return answers.shift()!;
        };
    });

    afterEach(() => {
        window.prompt = global.prompt = prompt;
    });

    const renameItem = async () => {
        let items: Jspreadsheet.ContextMenuItem[] | undefined;
        const worksheet = await load({
            worksheets: [{ data: [['a', 'b']], columns: [{ title: 'Name' }, { title: 'Age' }] }],
            contextMenu: (instance, x, y, e, defaultItems) => {
                items = defaultItems;
                return false;
            },
        });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        return { worksheet, item: items!.find((item) => item.title === 'Rename this column') as Jspreadsheet.ContextMenuOption };
    };

    it('keeps the title when the prompt of "Rename this column" is cancelled', async () => {
        const { worksheet, item } = await renameItem();

        answers.push(null);
        // @ts-expect-error -- the test calls the item without the element and the event, which the item does not read
        item.onclick!();

        expect(questions).to.eql([['Column name', 'Name']]);
        expect(worksheet.getHeader(0)).to.equal('Name');
    });

    it('sets the title that the prompt of "Rename this column" gives', async () => {
        const { worksheet, item } = await renameItem();

        answers.push('Surname');
        // @ts-expect-error -- the test calls the item without the element and the event, which the item does not read
        item.onclick!();

        expect(worksheet.getHeader(0)).to.equal('Surname');
    });

    it('asks for the title when the press on the selected header is held', async () => {
        const worksheet = await load({
            worksheets: [{ data: [['a', 'b']], columns: [{ title: 'Name' }, { title: 'Age' }], columnResize: false, columnDrag: false }],
        });

        mouse(worksheet.headers[1], 'mousedown');
        mouse(document, 'mouseup');

        answers.push(null);
        mouse(worksheet.headers[1], 'mousedown');
        await new Promise((resolve) => setTimeout(resolve, 850));
        mouse(document, 'mouseup');

        expect(questions).to.eql([['Column name', 'Age']]);
        expect(worksheet.getHeader(1)).to.equal('Age');
    });
});
