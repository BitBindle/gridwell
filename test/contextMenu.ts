import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('The default context menu', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Jspreadsheet.SpreadsheetOptions) =>
        new Promise<Worksheet[]>((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance) });
        });

    const contextmenu = (element: Element) => element.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

    // The items of the menu that jSuites shows, after its header
    const menuItems = () => [...root.querySelectorAll<HTMLElement>('.jss_contextmenu > div:not(.header)')];

    const titles = () => menuItems().map((item) => item.querySelector('a')!.textContent);

    const click = (title: string) => {
        const item = menuItems().find((element) => element.querySelector('a')!.textContent === title)!;
        item.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }));
    };

    let alert: typeof window.alert;
    let prompt: typeof window.prompt;
    let execCommand: typeof document.execCommand;
    let messages: (string | [string | undefined, string | undefined])[];
    let answers: string[];

    beforeEach(() => {
        alert = window.alert;
        prompt = window.prompt;
        execCommand = document.execCommand;
        messages = [];
        answers = [];
        // jsdom has no execCommand, which a copy calls
        document.execCommand = () => true;
        window.alert = global.alert = (text: string) => messages.push(text);
        window.prompt = global.prompt = (text?: string, value?: string) => {
            messages.push([text, value]);
            return answers.shift() ?? null;
        };
    });

    afterEach(() => {
        window.alert = global.alert = alert;
        window.prompt = global.prompt = prompt;
        document.execCommand = execCommand;
    });

    describe('on a header', () => {
        const data = () => [
            ['b', 2, 'x'],
            ['a', 1, 'y'],
        ];

        it('shows the items of a column', async () => {
            const [worksheet] = await load({ worksheets: [{ data: data() }] });

            contextmenu(worksheet.headers[1]);

            expect(root.querySelector('.jss_contextmenu')!.classList.contains('jcontextmenu-focus')).to.equal(true);
            expect(titles()).to.eql([
                'Insert a new column before',
                'Insert a new column after',
                'Delete selected columns',
                'Rename this column',
                'Order ascending',
                'Order descending',
                'Copy...',
                'Save as...',
                'About',
            ]);
        });

        it('inserts a column before and after the column', async () => {
            const [worksheet] = await load({ worksheets: [{ data: data() }] });

            contextmenu(worksheet.headers[1]);
            click('Insert a new column before');

            expect(worksheet.getData()[0]).to.eql(['b', '', 2, 'x']);

            contextmenu(worksheet.headers[0]);
            click('Insert a new column after');

            expect(worksheet.getData()[0]).to.eql(['b', '', '', 2, 'x']);
        });

        it('deletes the selected columns', async () => {
            const [worksheet] = await load({ worksheets: [{ data: data() }] });

            contextmenu(worksheet.headers[1]);
            click('Delete selected columns');

            expect(worksheet.getData()).to.eql([
                ['b', 'x'],
                ['a', 'y'],
            ]);
        });

        it('orders the rows by the column', async () => {
            const [worksheet] = await load({ worksheets: [{ data: data() }] });

            contextmenu(worksheet.headers[1]);
            click('Order ascending');

            expect(worksheet.getColumnData(0)).to.eql(['a', 'b']);

            contextmenu(worksheet.headers[1]);
            click('Order descending');

            expect(worksheet.getColumnData(0)).to.eql(['b', 'a']);
        });

        it('leaves out the items that the options turn off', async () => {
            const [worksheet] = await load({
                worksheets: [{ data: data(), allowInsertColumn: false, allowDeleteColumn: false, allowRenameColumn: false, columnSorting: false }],
                allowExport: false,
                about: false,
            });

            contextmenu(worksheet.headers[1]);

            expect(titles()).to.eql(['Copy...']);
        });
    });

    describe('on a row number', () => {
        it('inserts a row before and after the row and deletes the selected rows', async () => {
            const [worksheet] = await load({ worksheets: [{ data: [['a'], ['b']] }] });

            contextmenu(worksheet.rows[1].element.children[0]);

            expect(titles()).to.eql(['Insert a new row before', 'Insert a new row after', 'Delete selected rows', 'Copy...', 'Save as...', 'About']);

            click('Insert a new row before');
            expect(worksheet.getColumnData(0)).to.eql(['a', '', 'b']);

            contextmenu(worksheet.rows[2].element.children[0]);
            click('Insert a new row after');
            expect(worksheet.getColumnData(0)).to.eql(['a', '', 'b', '']);

            contextmenu(worksheet.rows[0].element.children[0]);
            click('Delete selected rows');
            expect(worksheet.getColumnData(0)).to.eql(['', 'b', '']);
        });

        it('leaves out the row items that the options turn off', async () => {
            const [worksheet] = await load({ worksheets: [{ data: [['a'], ['b']], allowInsertRow: false, allowDeleteRow: false }], about: false });

            contextmenu(worksheet.rows[1].element.children[0]);

            expect(titles()).to.eql(['Copy...', 'Save as...']);
        });
    });

    describe('on a cell', () => {
        it('adds, edits and clears the comments of the cell', async () => {
            const [worksheet] = await load({ worksheets: [{ data: [['a', 'b']] }] });

            contextmenu(worksheet.records[0][1].element);
            expect(titles()).to.include('Add comments');

            answers.push('Hello');
            click('Add comments');

            expect(messages).to.eql([['Comments', '']]);
            expect(worksheet.getComments('B1')).to.equal('Hello');

            contextmenu(worksheet.records[0][1].element);
            expect(titles()).to.include('Edit comments');

            answers.push('');
            click('Edit comments');
            expect(worksheet.getComments('B1')).to.equal('Hello');

            contextmenu(worksheet.records[0][1].element);
            click('Clear comments');
            expect(worksheet.getComments('B1')).to.equal('');
        });

        it('has no comment items when the comments are turned off', async () => {
            const [worksheet] = await load({ worksheets: [{ data: [['a', 'b']], allowComments: false }] });

            contextmenu(worksheet.records[0][1].element);

            expect(titles()).to.not.include('Add comments');
        });

        it('copies the selected cells', async () => {
            const [worksheet] = await load({ worksheets: [{ data: [['a', 'b']] }] });

            worksheet.updateSelectionFromCoords(0, 0, 1, 0);
            contextmenu(worksheet.records[0][1].element);
            click('Copy...');

            expect(worksheet.clipboard.data).to.equal('a\tb');
        });

        it('pastes the text of the clipboard into the selection', async () => {
            // The bundle reads the navigator of the global object, which Node defines with a getter
            const descriptor = Object.getOwnPropertyDescriptor(global, 'navigator');
            Object.defineProperty(global, 'navigator', { configurable: true, value: { clipboard: { readText: () => Promise.resolve('x\ty') } } });

            try {
                const [worksheet] = await load({ worksheets: [{ data: [['a', 'b', 'c']] }] });

                contextmenu(worksheet.records[0][1].element);
                click('Paste...');
                await new Promise((resolve) => setTimeout(resolve));

                expect(worksheet.getData()).to.eql([['a', 'x', 'y']]);
            } finally {
                Object.defineProperty(global, 'navigator', descriptor!);
            }
        });
    });

    it('downloads the worksheet with "Save as"', async () => {
        const [worksheet] = await load({ worksheets: [{ data: [['a']] }] });
        let downloads = 0;
        worksheet.download = () => downloads++;

        contextmenu(worksheet.records[0][0].element);
        click('Save as...');

        expect(downloads).to.equal(1);
    });

    it('shows the version with "About"', async () => {
        const [worksheet] = await load({ worksheets: [{ data: [['a']] }] });

        contextmenu(worksheet.records[0][0].element);
        click('About');

        expect(messages).to.have.length(1);
        expect(messages[0]).to.match(/Gridwell/);
    });

    it('shows the text of the about option with "About"', async () => {
        const [worksheet] = await load({ worksheets: [{ data: [['a']] }], about: 'My application' });

        contextmenu(worksheet.records[0][0].element);
        click('About');

        expect(messages).to.eql(['My application']);
    });

    it('selects the columns of a nested header', async () => {
        let call: [number | string | null, number | string | null, string, (string | undefined)[]] | undefined;
        const [worksheet] = await load({
            worksheets: [
                {
                    minDimensions: [3, 2],
                    nestedHeaders: [
                        [
                            { title: 'First', colspan: 2 },
                            { title: 'Second', colspan: 1 },
                        ],
                    ],
                },
            ],
            contextMenu: (instance, x, y, e, items, role) => {
                call = [x, y, role, items.map((item) => item.title)];
            },
        });

        contextmenu(root.querySelector('.jss_nested td[data-column="0,1"]')!);

        expect(worksheet.selection.range).to.eql([0, 0, 1, 1]);
        expect(call![0]).to.equal(0);
        expect(call![2]).to.equal('nested');
        expect(call![3]).to.eql(['Save as...', 'About']);
    });

    it('selects all cells from the corner of the headers', async () => {
        let role: string | undefined;
        const [worksheet] = await load({
            worksheets: [{ minDimensions: [2, 2] }],
            contextMenu: (instance, x, y, e, items, r) => {
                role = r;
                return false;
            },
        });

        contextmenu(root.querySelector('.jss_selectall')!);

        expect(role).to.equal('select-all');
        expect(worksheet.selection.range).to.eql([0, 0, 1, 1]);
    });

    it('passes the index of a footer cell', async () => {
        let call: [number | string | null, number | string | null, string] | undefined;
        await load({
            worksheets: [{ minDimensions: [3, 2], footers: [['a', 'b', 'c']] }],
            contextMenu: (instance, x, y, e, items, role) => {
                call = [x, y, role];
                return false;
            },
        });

        contextmenu(root.querySelector('tfoot tr td:nth-child(3)')!);

        expect(call).to.eql([1, 0, 'footer']);
    });

    it('passes the index of a tab', async () => {
        let call: [number | string | null, string] | undefined;
        await load({
            tabs: true,
            worksheets: [{ minDimensions: [1, 1] }, { minDimensions: [1, 1] }],
            contextMenu: (instance, x, y, e, items, role) => {
                call = [x, role];
                return false;
            },
        });

        contextmenu(root.querySelectorAll('.jtabs-headers > button')[1]);

        expect(call).to.eql([1, 'tabs']);
    });

    it('shows the items that the contextMenu option returns', async () => {
        let clicks = 0;
        const [worksheet] = await load({
            worksheets: [{ minDimensions: [1, 1] }],
            contextMenu: () => [{ title: 'Mine', onclick: () => clicks++ }],
        });

        contextmenu(worksheet.records[0][0].element);
        expect(titles()).to.eql(['Mine']);

        click('Mine');
        expect(clicks).to.equal(1);
    });

    it('lets a plugin change the items', async () => {
        const calls: [number | null, number | null, string][] = [];
        const [worksheet] = await load({
            worksheets: [{ minDimensions: [2, 2] }],
            plugins: {
                first: () => ({
                    contextMenu(instance, x, y, e, items, role) {
                        calls.push([x, y, role]);
                        return [...items, { title: 'Plugin' }];
                    },
                }),
                second: () => ({
                    contextMenu() {},
                }),
                third: () => ({}),
            },
        });

        contextmenu(worksheet.records[1][0].element);
        expect(calls).to.eql([[0, 1, 'cell']]);
        expect(titles().at(-1)).to.equal('Plugin');

        contextmenu(worksheet.headers[1]);
        expect(calls[1]).to.eql([1, null, 'header']);
    });

    it('does not open while a cell is being edited', async () => {
        const [worksheet] = await load({ worksheets: [{ data: [['a', 'b']] }] });

        worksheet.openEditor(worksheet.records[0][0].element);
        const event = new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true });
        worksheet.records[0][1].element.dispatchEvent(event);

        expect(event.defaultPrevented).to.equal(true);
        expect(menuItems()).to.have.length(0);
    });
});

describe('The context menu', () => {
    const load = (options: Jspreadsheet.SpreadsheetOptions, element: HTMLElement = root) =>
        new Promise<Worksheet[]>((resolve) => {
            const instance = jspreadsheet(element, { ...options, onload: () => resolve(instance) });
        });

    const menu = (element: Element = root) => element.querySelector<HTMLElement>('.jss_contextmenu')!;
    const items = (element: Element = menu()) => [...element.children].filter((child) => child.getAttribute('role') === 'menuitem') as HTMLElement[];
    const key = (element: Element, name: string) => {
        const event = new window.KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true });
        element.dispatchEvent(event);
        return event;
    };
    const open = (element: Element, init: MouseEventInit = {}) =>
        element.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, ...init }));

    const options = (choices: string[]): Jspreadsheet.SpreadsheetOptions => ({
        worksheets: [{ minDimensions: [2, 2] }],
        contextMenu: () => [
            {
                title: 'First',
                shortcut: 'Ctrl + F',
                tooltip: 'The first item',
                icon: 'star',
                id: 'first',
                onclick: (element, e) => choices.push(`First ${e.type}`),
            },
            { type: 'line' },
            { title: 'Disabled', disabled: true, onclick: () => choices.push('Disabled') },
            {
                title: 'More',
                submenu: [
                    { title: 'Inner', onclick: () => choices.push('Inner') },
                    { title: 'Other', onclick: () => choices.push('Other') },
                ],
            },
            { title: 'Last', onclick: () => choices.push('Last') },
        ],
    });

    it('is a menu of items and lines, with the attributes of the items', async () => {
        const [worksheet] = await load(options([]));

        open(worksheet.records[0][0].element);

        expect(menu().getAttribute('role')).to.equal('menu');
        expect(menu().getAttribute('aria-label')).to.equal('Menu');
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(true);
        expect(menu().querySelector('.header')!.getAttribute('aria-hidden')).to.equal('true');
        expect(menu().querySelector(':scope > hr')!.getAttribute('role')).to.equal('separator');
        expect(items().map((item) => item.querySelector('a')!.textContent)).to.eql(['First', 'Disabled', 'More', 'Last']);
        expect(items().every((item) => item.getAttribute('tabindex') === '-1')).to.equal(true);

        const [first, disabled, more] = items();
        expect([first.title, first.getAttribute('data-icon'), first.id, first.querySelector('span')!.textContent]).to.eql([
            'The first item',
            'star',
            'first',
            'Ctrl + F',
        ]);
        expect(disabled.getAttribute('aria-disabled')).to.equal('true');
        expect(disabled.classList.contains('jcontextmenu-disabled')).to.equal(true);
        expect(more.getAttribute('aria-haspopup')).to.equal('menu');
        expect(more.getAttribute('aria-expanded')).to.equal('false');
        expect(more.querySelector('[role="menu"]')!.getAttribute('aria-label')).to.equal('More');
    });

    it('shows the titles as text', async () => {
        const [worksheet] = await load({ worksheets: [{ minDimensions: [1, 1] }], contextMenu: () => [{ title: '<b>Bold</b>' }] });

        open(worksheet.records[0][0].element);

        expect(items()[0].querySelector('a')!.textContent).to.equal('<b>Bold</b>');
        expect(menu().querySelector('b')).to.equal(null);
    });

    it('moves the focus to the first item, and between the items with the arrows, Home and End', async () => {
        const [worksheet] = await load(options([]));

        open(worksheet.records[0][0].element);
        const [first, disabled, more, last] = items();
        expect(document.activeElement).to.equal(first);

        key(first, 'ArrowDown');
        expect(document.activeElement).to.equal(disabled);
        key(disabled, 'ArrowDown');
        key(more, 'ArrowDown');
        expect(document.activeElement).to.equal(last);
        key(last, 'ArrowDown');
        expect(document.activeElement).to.equal(first);
        key(first, 'ArrowUp');
        expect(document.activeElement).to.equal(last);
        key(last, 'Home');
        expect(document.activeElement).to.equal(first);
        key(first, 'End');
        expect(document.activeElement).to.equal(last);

        // The menu itself has the focus when it has no item
        key(menu(), 'ArrowDown');
        expect(document.activeElement).to.equal(first);
        key(menu(), 'ArrowUp');
        expect(document.activeElement).to.equal(last);
    });

    it('chooses an item with Enter or Space, closes and gives the focus back to the cell', async () => {
        const choices: string[] = [];
        const [worksheet] = await load(options(choices));
        const cell = worksheet.records[0][0].element;
        cell.focus();

        open(cell);
        const event = key(items()[0], 'Enter');

        expect(choices).to.eql(['First keydown']);
        expect(event.defaultPrevented).to.equal(true);
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
        expect(document.activeElement).to.equal(cell);

        open(cell);
        key(items()[3], ' ');
        expect(choices).to.eql(['First keydown', 'Last']);
    });

    it('does nothing for a disabled item', async () => {
        const choices: string[] = [];
        const [worksheet] = await load(options(choices));

        open(worksheet.records[0][0].element);
        const disabled = items()[1];
        key(disabled, 'Enter');
        disabled.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }));

        expect(choices).to.eql([]);
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(true);
    });

    it('opens a submenu with the right arrow, closes it with the left arrow or Escape, and chooses its items', async () => {
        const choices: string[] = [];
        const [worksheet] = await load(options(choices));

        open(worksheet.records[0][0].element);
        const more = items()[2];
        const [inner, other] = items(more.querySelector('[role="menu"]')!);

        key(more, 'ArrowRight');
        expect(more.getAttribute('aria-expanded')).to.equal('true');
        expect(more.classList.contains('jcontextmenu-expanded')).to.equal(true);
        expect(document.activeElement).to.equal(inner);

        key(inner, 'ArrowDown');
        expect(document.activeElement).to.equal(other);

        key(other, 'ArrowLeft');
        expect(more.getAttribute('aria-expanded')).to.equal('false');
        expect(document.activeElement).to.equal(more);

        key(more, 'Enter');
        key(inner, 'Escape');
        expect(document.activeElement).to.equal(more);
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(true);

        key(more, 'ArrowRight');
        inner.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }));
        expect(choices).to.eql(['Inner']);
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
    });

    it('focuses and expands the item under the mouse, and chooses an item with the left button only', async () => {
        const choices: string[] = [];
        const [worksheet] = await load(options(choices));

        open(worksheet.records[0][0].element);
        const [first, , more] = items();

        more.dispatchEvent(new window.MouseEvent('mouseenter'));
        expect(document.activeElement).to.equal(more);
        expect(more.getAttribute('aria-expanded')).to.equal('true');
        more.dispatchEvent(new window.MouseEvent('mouseleave'));
        expect(more.getAttribute('aria-expanded')).to.equal('false');

        const press = new window.MouseEvent('mousedown', { bubbles: true, cancelable: true });
        first.dispatchEvent(press);
        expect(press.defaultPrevented).to.equal(true);

        first.querySelector('a')!.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true, button: 2 }));
        expect(choices).to.eql([]);
        first.querySelector('a')!.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }));
        expect(choices).to.eql(['First mouseup']);
    });

    it('closes with Escape or Tab, and the keys in the menu do not move the selection', async () => {
        const [worksheet] = await load(options([]));
        const cell = worksheet.records[0][0].element;
        worksheet.updateSelectionFromCoords(0, 0);
        cell.focus();

        open(cell);
        key(items()[0], 'ArrowDown');
        expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);

        key(document.activeElement!, 'Escape');
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
        expect(document.activeElement).to.equal(cell);

        open(cell);
        const tab = key(items()[0], 'Tab');
        expect(tab.defaultPrevented).to.equal(true);
        expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
        expect(document.activeElement).to.equal(cell);

        // Other keys are left to the browser
        open(cell);
        expect(key(items()[0], 'a').defaultPrevented).to.equal(false);
        expect(key(items()[0], 'ArrowLeft').defaultPrevented).to.equal(false);
    });

    it('closes on a press outside it, which keeps the focus where the press puts it, and on the wheel outside it', async () => {
        const [worksheet] = await load(options([]));
        const outside = document.createElement('button');
        document.body.appendChild(outside);

        try {
            open(worksheet.records[0][0].element);
            items()[0].dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
            menu().dispatchEvent(new window.WheelEvent('wheel', { bubbles: true }));
            expect(menu().classList.contains('jcontextmenu-focus')).to.equal(true);

            outside.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
            expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
            expect(document.activeElement).to.not.equal(worksheet.records[0][0].element);

            open(worksheet.records[0][0].element);
            outside.dispatchEvent(new window.WheelEvent('wheel', { bubbles: true }));
            expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
        } finally {
            outside.remove();
        }
    });

    it('closes the menu of another spreadsheet when it opens', async () => {
        const other = document.createElement('div');
        document.body.appendChild(other);

        try {
            const [worksheet] = await load(options([]));
            const [second] = await load(options([]), other);

            open(worksheet.records[0][0].element);
            open(second.records[0][0].element);

            expect(menu().classList.contains('jcontextmenu-focus')).to.equal(false);
            expect(menu(other).classList.contains('jcontextmenu-focus')).to.equal(true);
        } finally {
            jspreadsheet.destroy(other as Jspreadsheet.JspreadsheetInstanceElement);
            other.remove();
        }
    });

    it('opens at the pointer, and under the element when a key opens it', async () => {
        const [worksheet] = await load(options([]));
        const cell = worksheet.records[1][1].element;

        open(cell, { clientX: 100, clientY: 50, button: 2 });
        expect([menu().style.left, menu().style.top]).to.eql(['100px', '50px']);

        cell.getBoundingClientRect = () => ({ left: 30, bottom: 40, top: 20, right: 90, width: 60, height: 20, x: 30, y: 20, toJSON() {} });
        open(cell);
        expect([menu().style.left, menu().style.top]).to.eql(['30px', '40px']);

        worksheet.parent.contextMenu.contextmenu.open({ x: 5, y: 6 });
        expect([menu().style.left, menu().style.top]).to.eql(['5px', '6px']);
    });
});
