import { expect } from 'chai';

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
});
