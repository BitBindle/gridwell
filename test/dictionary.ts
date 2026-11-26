import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

describe('Translate the texts with setDictionary', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (columns?: { type: 'calendar' }[]) =>
        new Promise<Worksheet>((resolve) => {
            const instance = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2], columns }], onload: () => resolve(instance[0]) });
        });

    const titles = () => [...root.querySelectorAll('.jss_contextmenu > div:not(.header) a')].map((item) => item.textContent);

    afterEach(() => {
        // An empty translation shows the English text again
        jspreadsheet.setDictionary({ 'Insert a new column before': '', 'Delete selected columns': '', Update: '' });
    });

    it('translates the items of the context menu', async () => {
        const worksheet = await load();
        jspreadsheet.setDictionary({ 'Insert a new column before': 'Neue Spalte davor einfügen' });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(titles()[0]).to.equal('Neue Spalte davor einfügen');
        expect(titles()[1]).to.equal('Insert a new column after');
    });

    it('keeps the translations of earlier calls', async () => {
        const worksheet = await load();
        jspreadsheet.setDictionary({ 'Insert a new column before': 'Neue Spalte davor einfügen' });
        jspreadsheet.setDictionary({ 'Delete selected columns': 'Ausgewählte Spalten löschen' });

        worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(titles()).to.include.members(['Neue Spalte davor einfügen', 'Ausgewählte Spalten löschen']);
    });

    it('translates the texts of the calendar', async () => {
        const worksheet = await load([{ type: 'calendar' }]);
        jspreadsheet.setDictionary({ Update: 'Übernehmen' });

        worksheet.openEditor(worksheet.records[0][0].element);

        const buttons = [...document.querySelectorAll('.jcalendar-controls button')].map((button) => button.textContent);
        expect(buttons).to.include('Übernehmen');
        worksheet.closeEditor(worksheet.records[0][0].element, false);
    });
});
