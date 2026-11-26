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
            // The bundle reads the navigator of the global object, which jsdom-global defines with a getter
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
        expect(messages[0]).to.match(/Jspreadsheet CE/);
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
