import { expect } from 'chai';

import gridwell, { type Gridwell, type Worksheet } from './support/gridwell';

describe('Touch', () => {
    it('opens no editor after a long touch when the worksheet is not editable', async () => {
        const events: string[] = [];
        const worksheet = await new Promise<Worksheet>((resolve) => {
            const instance = gridwell(root, {
                worksheets: [{ minDimensions: [3, 3], editable: false }],
                oneditionstart: () => events.push('start'),
                onload: () => resolve(instance[0]),
            });
        });

        worksheet.records[1][1].element.dispatchEvent(new window.Event('touchstart', { bubbles: true }));

        // A long touch opens the editor after 500 ms
        await new Promise((resolve) => setTimeout(resolve, 600));

        expect(worksheet.editor.cell).to.not.be.ok;
        expect(events).to.eql([]);
        expect(worksheet.selection.range!.map(Number)).to.eql([1, 1, 1, 1]);
    });

    const load = (element: HTMLElement, options: Gridwell.SpreadsheetOptions) =>
        new Promise<Worksheet>((resolve) => {
            const instance = gridwell(element, { ...options, onload: () => resolve(instance[0]) });
        });

    const touch = (element: Element, type: string) => element.dispatchEvent(new window.Event(type, { bubbles: true }));

    it('opens the editor after a long touch on a cell', async () => {
        const worksheet = await load(root, { worksheets: [{ data: [['a', 'b']] }] });

        touch(worksheet.records[0][1].element, 'touchstart');
        await new Promise((resolve) => setTimeout(resolve, 600));
        touch(worksheet.records[0][1].element, 'touchend');

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(worksheet.editor.x).to.equal(1);
    });

    it('opens no editor when the touch ends early', async () => {
        const worksheet = await load(root, { worksheets: [{ data: [['a', 'b']] }] });

        touch(worksheet.records[0][0].element, 'touchstart');
        touch(worksheet.records[0][0].element, 'touchend');
        await new Promise((resolve) => setTimeout(resolve, 600));

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
    });

    it('moves the active worksheet to the touched spreadsheet, and away on a touch outside', async () => {
        const other = document.createElement('div');
        document.body.appendChild(other);

        try {
            const first = await load(root, { worksheets: [{ minDimensions: [2, 2] }] });
            const second = await load(other, { worksheets: [{ minDimensions: [2, 2] }] });

            touch(first.records[0][0].element, 'touchstart');
            touch(first.records[0][0].element, 'touchend');
            expect(gridwell.current).to.equal(first);

            touch(second.records[1][1].element, 'touchstart');
            touch(second.records[1][1].element, 'touchend');
            expect(gridwell.current).to.equal(second);
            expect(first.selection.range).to.equal(null);

            touch(document.body, 'touchstart');
            expect(gridwell.current).to.equal(null);
            expect(second.selection.range).to.equal(null);
        } finally {
            // gridwell() made the element the element of a spreadsheet
            gridwell.destroy(other as Gridwell.GridwellInstanceElement);
            other.remove();
        }
    });
});
