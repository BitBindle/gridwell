import { expect } from 'chai';

import gridwell, { type Worksheet } from './support/gridwell';

describe('Translate the texts with setDictionary', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (columns?: { type: 'calendar' }[]) =>
        new Promise<Worksheet>((resolve) => {
            const instance = gridwell(root, { worksheets: [{ minDimensions: [2, 2], columns }], onload: () => resolve(instance[0]) });
        });

    const titles = () => [...root.querySelectorAll('.jss_contextmenu > div:not(.header) a')].map((item) => item.textContent);

    afterEach(() => {
        // An empty translation shows the English text again
        gridwell.setDictionary({ 'Insert a new column before': '', 'Delete selected columns': '', Update: '', '(Blanks)': '' });
    });

    it('translates the items of the context menu', async () => {
        const worksheet = await load();
        gridwell.setDictionary({ 'Insert a new column before': 'Neue Spalte davor einfügen' });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(titles()[0]).to.equal('Neue Spalte davor einfügen');
        expect(titles()[1]).to.equal('Insert a new column after');
    });

    it('keeps the translations of earlier calls', async () => {
        const worksheet = await load();
        gridwell.setDictionary({ 'Insert a new column before': 'Neue Spalte davor einfügen' });
        gridwell.setDictionary({ 'Delete selected columns': 'Ausgewählte Spalten löschen' });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(titles()).to.include.members(['Neue Spalte davor einfügen', 'Ausgewählte Spalten löschen']);
    });

    it('translates the texts of the calendar', async () => {
        const worksheet = await load([{ type: 'calendar' }]);
        gridwell.setDictionary({ Update: 'Übernehmen' });

        worksheet.openEditor(worksheet.records[0][0].element);

        const buttons = [...document.querySelectorAll('.jcalendar-controls button')].map((button) => button.textContent);
        expect(buttons).to.include('Übernehmen');
        worksheet.closeEditor(worksheet.records[0][0].element, false);
    });

    it('translates the item of the empty cells in the filter of a column (issue #150)', async () => {
        const worksheet = gridwell(root, { worksheets: [{ filters: true, data: [['a'], ['']] }] })[0];
        gridwell.setDictionary({ '(Blanks)': '(Leer)' });

        worksheet.openFilter(0);

        const items = [...worksheet.filter!.querySelectorAll('.jdropdown-item .jdropdown-description')].map((item) => item.textContent);
        expect(items).to.eql(['a', '(Leer)']);
    });
});
