import { expect } from 'chai';

import { stubFetch } from './support/fetch';
import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// The constructor builds the worksheets after it returns. Resolve when all worksheets are ready.
function create(options?: Partial<Jspreadsheet.SpreadsheetOptions>) {
    return new Promise<Worksheet[]>((resolve) => {
        const worksheets = jspreadsheet(
            root,
            Object.assign<Jspreadsheet.SpreadsheetOptions, typeof options>(
                {
                    worksheets: [
                        { worksheetName: 'One', data: [['1', '2']] },
                        { worksheetName: 'Two', data: [['3', '4']], freezeRows: 1 },
                    ],
                    onload: () => resolve(worksheets),
                },
                options
            )
        );
    });
}

function tabNames() {
    return Array.from(root.querySelectorAll('.jtabs-headers > :not(.jtabs-border)')).map((tab) => tab.textContent);
}

describe('loadWorksheet', () => {
    it('builds the worksheet again in its tab, and keeps the worksheet object, its name and the open worksheet', async () => {
        const worksheets = await create();
        const [one, two] = worksheets;
        two.history.ignore = true;
        jspreadsheet.current = one;

        const options = {
            data: [
                ['x', 'y'],
                ['z', 'w'],
            ],
        };
        const loaded = await one.loadWorksheet(1, options);

        expect(loaded).to.equal(two);
        expect(worksheets[1]).to.equal(two);
        expect(two.getData()).to.eql([
            ['x', 'y'],
            ['z', 'w'],
        ]);
        expect(two.options).to.equal(options);
        expect(two.options.worksheetName).to.equal('Two');
        expect(two.table.getAttribute('aria-label')).to.equal('Two');
        expect(two.history.ignore).to.equal(true);
        expect(two.element.querySelectorAll('table').length).to.equal(1);
        expect(two.element.querySelectorAll('.jss_frozen_row').length).to.equal(0);
        expect(one.getWorksheetActive()).to.equal(0);
        expect(jspreadsheet.current).to.equal(one);
        expect(one.parent.config.worksheets[1]).to.equal(options);
        expect(tabNames()).to.eql(['One', 'Two']);
    });

    it('gives the worksheet a new name that no other worksheet has, and keeps the name otherwise', async () => {
        const [one] = await create();

        await one.loadWorksheet(1, { worksheetName: 'New', data: [['a']] });
        expect(tabNames()).to.eql(['One', 'New']);

        const two = await one.loadWorksheet(1, { worksheetName: 'One', data: [['b']] });
        expect(two!.options.worksheetName).to.equal('New');
        expect(tabNames()).to.eql(['One', 'New']);
    });

    it('closes the open editor without a save, and starts a new history', async () => {
        const [one] = await create();
        one.setValue('A1', 'changed');
        const cell = one.getCell('A1');
        one.openEditor(cell);
        (cell.children[0] as HTMLInputElement).value = 'typed';

        await one.loadWorksheet(0, { data: [['new']] });

        expect(one.editor.isOpen()).to.equal(false);
        expect(one.getValue('A1')).to.equal('new');
        expect(one.history.index).to.equal(-1);
        one.undo();
        expect(one.getValue('A1')).to.equal('new');
    });

    it('moves the focus from the old table to the tab stop of the new table', async () => {
        const [one] = await create();
        one.updateSelectionFromCoords(1, 0, 1, 0);
        one.getCell('B1').focus();

        await one.loadWorksheet(0, { data: [['a', 'b']] });

        expect(document.activeElement).to.equal(one.getCell('A1'));
    });

    it('gives undefined for a position without a worksheet', async () => {
        const [one] = await create();

        expect(await one.loadWorksheet(5, { data: [['a']] })).to.equal(undefined);
    });

    it('selects the tab stop of the new table when it has the focus, so that the keys move the selection', async () => {
        const [one] = await create();
        one.updateSelectionFromCoords(1, 0, 1, 0);
        one.getCell('B1').focus();

        await one.loadWorksheet(0, {
            data: [
                ['a', 'b'],
                ['c', 'd'],
            ],
        });
        document.activeElement!.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, key: 'ArrowDown', keyCode: 40, which: 40 }));

        expect(one.getSelection()).to.eql([0, 1, 0, 1]);
    });

    it('keeps the active worksheet while the data of a url loads, and uses the data of the last of two loads', async () => {
        const server = stubFetch();
        try {
            const worksheets = await create();
            const [one, two] = worksheets;
            jspreadsheet.current = one;
            server.responses['/rows'] = { body: [['u1', 'u2']] };

            const first = one.loadWorksheet(1, { url: '/rows' });
            expect(jspreadsheet.current).to.equal(one);
            const second = one.loadWorksheet(1, { data: [['d1']] });

            expect(await first).to.equal(two);
            expect(await second).to.equal(two);
            await server.settle();
            expect(two.getData()).to.eql([['d1']]);
            expect(two.element.querySelectorAll('table').length).to.equal(1);
            expect(jspreadsheet.current).to.equal(one);
        } finally {
            server.restore();
        }
    });

    it('dispatches onrenameworksheet for a new name', async () => {
        const events: unknown[][] = [];
        const [one] = await create({
            onrenameworksheet: (worksheet, index, name, oldName) => {
                events.push([index, name, oldName]);
            },
        });

        await one.loadWorksheet(1, { worksheetName: 'New', data: [['a']] });

        expect(events).to.eql([[1, 'New', 'Two']]);
    });

    it('keeps the styles of the cells when it builds the worksheet again with its own options', async () => {
        const [one] = await create();
        one.setStyle('A1', 'color', 'red');

        await one.loadWorksheet(0, one.options);

        expect(one.getCell('A1').style.color).to.equal('red');
    });

    it('places no frozen rows of the old table after the load', async () => {
        const errors: unknown[] = [];
        const listener = (e: ErrorEvent) => errors.push(e.error);
        window.addEventListener('error', listener);
        const server = stubFetch();
        try {
            const [one, two] = await create();
            server.responses['/rows'] = { body: [['u1']] };

            two.insertRow(1, 0, true);
            await one.loadWorksheet(1, { url: '/rows', freezeRows: 1 });
            await server.settle();

            expect(errors).to.eql([]);
            expect(two.getData()).to.eql([['u1']]);
        } finally {
            server.restore();
            window.removeEventListener('error', listener);
        }
    });

    describe('while data loads', () => {
        let server: ReturnType<typeof stubFetch>;

        beforeEach(() => {
            server = stubFetch();
        });

        afterEach(() => {
            server.restore();
        });

        it('keeps a worksheet active that the user makes active during the load', async () => {
            const [one, two] = await create({
                worksheets: [
                    { worksheetName: 'One', data: [['1']] },
                    { worksheetName: 'Two', data: [['2']] },
                    { worksheetName: 'Three', data: [['3']] },
                ],
            });
            jspreadsheet.current = one;
            server.responses['/rows'] = { body: [['u']] };

            const loading = one.loadWorksheet(1, { url: '/rows' });
            const three = one.parent.worksheets[2];
            jspreadsheet.current = three;
            await loading;

            expect(jspreadsheet.current).to.equal(three);
            expect(two.getData()).to.eql([['u']]);
        });

        it('leaves the focus where the user put it during the load', async () => {
            const [one] = await create();
            one.updateSelectionFromCoords(0, 0, 0, 0);
            one.getCell('A1').focus();
            server.responses['/rows'] = { body: [['u']] };
            const input = document.createElement('input');
            root.appendChild(input);

            const loading = one.loadWorksheet(0, { url: '/rows' });
            input.focus();
            await loading;

            expect(document.activeElement).to.equal(input);
        });

        it('waits for the source of a dropdown column, and uses the options of the last of two loads', async () => {
            const [one, two] = await create();
            server.responses['/source'] = { body: ['x', 'y'] };

            const first = one.loadWorksheet(1, { data: [['a']], columns: [{ type: 'dropdown', url: '/source' }] });
            const second = one.loadWorksheet(1, { data: [['b']] });
            await first;
            await second;
            await server.settle();

            expect(two.element.querySelectorAll('table').length).to.equal(1);
            expect(two.getData()).to.eql([['b']]);

            await one.loadWorksheet(1, { data: [['x']], columns: [{ type: 'dropdown', url: '/source' }] });
            expect(two.getCell('A1').textContent).to.equal('x');
        });

        it('stops the timers of the old table', async () => {
            const errors: unknown[] = [];
            const listener = (e: ErrorEvent) => errors.push(e.error);
            window.addEventListener('error', listener);
            try {
                const data = Array.from({ length: 300 }, (row, y) => [`${y}`]);
                const [one] = await create({ worksheets: [{ worksheetName: 'One', data, lazyLoading: true, tableOverflow: true }] });
                server.responses['/rows'] = { body: [['u']] };

                // The wheel loads rows 100 ms after the event
                one.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
                await one.loadWorksheet(0, { url: '/rows' });
                await new Promise((resolve) => setTimeout(resolve, 150));

                expect(errors).to.eql([]);
                expect(one.getData()).to.eql([['u']]);
            } finally {
                window.removeEventListener('error', listener);
            }
        });

        it('dispatches onrenameworksheet for a load that a later load replaces', async () => {
            const events: unknown[][] = [];
            const [one] = await create({
                onrenameworksheet: (worksheet, index, name, oldName) => {
                    events.push([index, name, oldName]);
                },
            });
            server.responses['/rows'] = { body: [['u']] };

            const first = one.loadWorksheet(1, { worksheetName: 'New', url: '/rows' });
            await one.loadWorksheet(1, { data: [['d']] });
            await first;

            expect(events).to.eql([[1, 'New', 'Two']]);
        });
    });

    it('closes an open filter without its events, and removes the class of the toolbar', async () => {
        const [one] = await create({
            toolbar: true,
            worksheets: [
                { worksheetName: 'One', data: [['a'], ['b']], filters: true },
                { worksheetName: 'Two', data: [['3', '4']] },
            ],
        });
        one.openFilter(0);
        expect(one.element.querySelector('.jdropdown-focus')).to.not.equal(null);

        await one.loadWorksheet(0, { data: [['x']], tableOverflow: true });

        expect(root.querySelector('.jdropdown-focus')).to.equal(null);
        expect(one.view.filters).to.eql([]);
        expect(one.element.classList.contains('with-toolbar')).to.equal(false);
    });
});
