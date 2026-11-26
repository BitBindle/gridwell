import { expect } from 'chai';

import gridwell, { type Gridwell, type Worksheet } from './support/gridwell';

describe('The mouse on the document', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options: Gridwell.SpreadsheetOptions) =>
        new Promise<Worksheet[]>((resolve) => {
            const instance = gridwell(root, { ...options, onload: () => resolve(instance) });
        });

    const mouse = (element: EventTarget, type: string, options?: MouseEventInit) =>
        element.dispatchEvent(new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...options }));

    const data = [
        ['a', 'b'],
        ['c', 'd'],
    ];

    const outside = () => {
        const element = document.createElement('div');
        document.body.appendChild(element);
        return element;
    };

    afterEach(() => {
        // A press outside the spreadsheets leaves no active worksheet for the next test
        mouse(document.body, 'mousedown');
        mouse(document.body, 'mouseup');
        for (const element of Array.from(document.body.children)) {
            if (element !== root) {
                element.remove();
            }
        }
    });

    it('opens the editor on a double click on the selected cell', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[0][1].element, 'mousedown');
        mouse(worksheet.records[0][1].element, 'mouseup');
        mouse(worksheet.records[0][1].element, 'dblclick');

        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(worksheet.editor.cell).to.equal(worksheet.records[0][1].element);
    });

    it('opens the editor on a double click on an element inside the selected cell', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });
        const span = document.createElement('span');
        worksheet.records[1][0].element.appendChild(span);

        mouse(worksheet.records[1][0].element, 'mousedown');
        mouse(worksheet.records[1][0].element, 'mouseup');
        mouse(span, 'dblclick');

        expect(worksheet.editor.cell).to.equal(worksheet.records[1][0].element);
    });

    it('does not open the editor on a double click on a cell that is not selected', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        worksheet.records[1][1].element.classList.remove('highlight');
        mouse(worksheet.records[1][1].element, 'dblclick');

        expect(worksheet.editor.isOpen()).to.equal(false);
    });

    it('does not open the editor on a double click when the worksheet is not editable', async () => {
        const [worksheet] = await load({ worksheets: [{ data, editable: false }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        mouse(worksheet.records[0][0].element, 'dblclick');

        expect(worksheet.editor.isOpen()).to.equal(false);
    });

    it('opens the filter of a column on a double click on it', async () => {
        const [worksheet] = await load({ worksheets: [{ data, filters: true }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        mouse(worksheet.filter!.children[2], 'dblclick');

        expect(worksheet.filter!.children[2].querySelector('.jdropdown')).to.not.equal(null);
        expect(worksheet.filter!.children[1].querySelector('.jdropdown')).to.equal(null);
    });

    it('selects all the cells on a press on the corner of the headers', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        mouse(worksheet.thead.querySelector('.jss_selectall')!, 'mousedown');
        mouse(document, 'mouseup');

        expect(worksheet.getSelection()).to.eql([0, 0, 1, 1]);
    });

    it('selects a cell in a browser that gives the button only as which', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        const event = new window.MouseEvent('mousedown', { bubbles: true, cancelable: true, buttons: 0, button: 0 });
        Object.defineProperty(event, 'which', { value: 1 });
        worksheet.records[1][1].element.dispatchEvent(event);
        mouse(document, 'mouseup', { buttons: 0 });

        expect(worksheet.getSelection()).to.eql([1, 1, 1, 1]);
    });

    it('does not select a cell on a press with the right button', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[1][1].element, 'mousedown', { buttons: 0, button: 2 });
        mouse(document, 'mouseup', { buttons: 0 });

        expect(worksheet.getSelection()).to.equal(null);
    });

    it('saves the open editor and clears the selection on a press outside the spreadsheet', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        worksheet.openEditor(worksheet.records[0][0].element);
        (worksheet.records[0][0].element.querySelector('input') as HTMLInputElement).value = 'x';
        mouse(outside(), 'mousedown');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValue('A1')).to.equal('x');
        expect(worksheet.getSelection()).to.equal(null);
    });

    it('keeps the selection on a press on a jss_object control outside the spreadsheet', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });
        const control = outside();
        control.classList.add('jss_object');

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][0].element, 'mouseup');
        mouse(control, 'mousedown');

        expect(worksheet.getSelection()).to.eql([0, 0, 0, 0]);
    });

    it('saves the open editor of a worksheet on a press on another worksheet', async () => {
        const [first, second] = await load({ worksheets: [{ data }, { data }] });

        mouse(first.records[0][0].element, 'mousedown');
        mouse(first.records[0][0].element, 'mouseup');
        first.openEditor(first.records[0][0].element);
        (first.records[0][0].element.querySelector('input') as HTMLInputElement).value = 'x';
        mouse(second.records[1][1].element, 'mousedown');
        mouse(second.records[1][1].element, 'mouseup');

        expect(first.editor.isOpen()).to.equal(false);
        expect(first.getValue('A1')).to.equal('x');
        expect(first.getSelection()).to.equal(null);
        expect(second.getSelection()).to.eql([1, 1, 1, 1]);
    });

    it('extends the selection while the button is down and not after it goes up', async () => {
        const [worksheet] = await load({ worksheets: [{ data }] });

        mouse(worksheet.records[0][0].element, 'mousedown');
        mouse(worksheet.records[0][1].element, 'mouseover');
        expect(worksheet.getSelection()).to.eql([0, 0, 1, 0]);

        // The button went up outside the window
        mouse(worksheet.records[1][1].element, 'mouseover', { buttons: 0 });
        expect(worksheet.getSelection()).to.eql([0, 0, 1, 0]);
    });

    it('does not extend the selection into another worksheet', async () => {
        const [first, second] = await load({ worksheets: [{ data }, { data }] });

        mouse(first.records[0][0].element, 'mousedown');
        mouse(second.records[1][1].element, 'mouseover');
        mouse(document, 'mouseup');

        expect(first.getSelection()).to.eql([0, 0, 0, 0]);
        expect(second.getSelection()).to.equal(null);
    });
});

describe('Typing in the search of a worksheet', () => {
    const key = (element: EventTarget, init: KeyboardEventInit) =>
        element.dispatchEvent(new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
    const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

    const data = [['Honda'], ['Mazda'], ['Honda']];

    it('searches the rows a moment after the last key', async () => {
        const [worksheet] = gridwell(root, { worksheets: [{ data, search: true }] });
        await wait(0);

        worksheet.searchInput.focus();
        worksheet.searchInput.value = 'Hon';
        key(worksheet.searchInput, { key: 'n' });
        expect(worksheet.view.results).to.equal(null);

        worksheet.searchInput.value = 'Honda';
        key(worksheet.searchInput, { key: 'a' });
        await wait(250);

        expect(worksheet.view.results).to.eql([0, 2]);
    });

    it('does not search when the mouse moves before the moment ends', async () => {
        const [worksheet] = gridwell(root, { worksheets: [{ data, search: true }] });
        await wait(0);

        worksheet.searchInput.focus();
        worksheet.searchInput.value = 'Mazda';
        key(worksheet.searchInput, { key: 'a' });
        worksheet.searchInput.dispatchEvent(new window.MouseEvent('mouseover', { bubbles: true }));
        await wait(250);

        expect(worksheet.view.results).to.equal(null);
    });
});
