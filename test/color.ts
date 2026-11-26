import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

// The key codes, which the keys of the editor read
const codes: Record<string, number> = { Enter: 13, Escape: 27, ' ': 32, Home: 36, End: 35, ArrowLeft: 37, ArrowUp: 38, ArrowRight: 39, ArrowDown: 40 };

const key = (element: Element, keyName: string) => {
    const keyCode = codes[keyName];
    const event = new KeyboardEvent('keydown', { key: keyName, keyCode, which: keyCode, bubbles: true, cancelable: true });
    element.dispatchEvent(event);
    return event;
};

const waitForListeners = () => new Promise((resolve) => setTimeout(resolve));

describe('The color picker', () => {
    describe('The editor of a color column', () => {
        const create = async (value = '#e57373', options?: object) => {
            const instance = jspreadsheet(root, {
                worksheets: [{ data: [[value, 'x']], columns: [{ type: 'color', options }, {}] }],
            });
            await waitForListeners();
            const worksheet = instance[0];
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            worksheet.openEditor(worksheet.getCell('A1'));
            const cell = worksheet.getCell('A1');
            return { worksheet, cell, picker: cell.querySelector<HTMLElement>('.jcolor')! };
        };

        it('is a dialog with a grid of named colors, and moves the focus to the color of the cell', async () => {
            const { picker } = await create();

            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            const dialog = picker.querySelector('.jcolor-content')!;
            expect(dialog.getAttribute('role')).to.equal('dialog');
            expect(dialog.getAttribute('aria-label')).to.equal('Color');

            const grid = picker.querySelector('table')!;
            expect(grid.getAttribute('role')).to.equal('grid');
            const cells = Array.from(grid.querySelectorAll('td'));
            expect(cells.length).to.equal(170);
            expect(cells.every((td) => td.getAttribute('role') === 'gridcell')).to.equal(true);

            const selected = grid.querySelector<HTMLElement>('[aria-selected="true"]')!;
            expect(selected.getAttribute('aria-label')).to.equal('#e57373');
            expect(selected.classList.contains('jcolor-selected')).to.equal(true);
            expect(document.activeElement).to.equal(selected);
            // The grid is one tab stop
            expect(cells.filter((td) => td.getAttribute('tabindex') === '0')).to.eql([selected]);

            // The buttons are buttons, and the other color has a name
            expect(Array.from(picker.querySelectorAll('button')).map((button) => button.textContent)).to.eql(['Reset', 'Done']);
            expect(picker.querySelector('.jcolor-custom')!.textContent).to.equal('Custom color');
            expect(picker.querySelector<HTMLInputElement>('.jcolor-custom input')!.value).to.equal('#e57373');
        });

        it('moves between the colors with the arrows, Home and End, and saves the color chosen with Enter', async () => {
            const { worksheet, picker } = await create('#ffebee');
            const rows = Array.from(picker.querySelectorAll('tr'));

            const right = key(document.activeElement!, 'ArrowRight');
            expect(right.defaultPrevented).to.equal(true);
            expect(document.activeElement).to.equal(rows[0].cells[1]);
            expect(rows[0].cells[1].getAttribute('tabindex')).to.equal('0');

            key(document.activeElement!, 'ArrowDown');
            expect(document.activeElement).to.equal(rows[1].cells[1]);
            key(document.activeElement!, 'End');
            expect(document.activeElement).to.equal(rows[1].cells[16]);
            key(document.activeElement!, 'ArrowRight');
            expect(document.activeElement).to.equal(rows[1].cells[16]);
            key(document.activeElement!, 'Home');
            expect(document.activeElement).to.equal(rows[1].cells[0]);
            key(document.activeElement!, 'ArrowUp');
            key(document.activeElement!, 'ArrowUp');
            expect(document.activeElement).to.equal(rows[0].cells[0]);
            key(document.activeElement!, 'ArrowDown');

            // Enter does not reach the grid, which would move the selection
            key(document.activeElement!, 'Enter');
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValue('A1')).to.equal('#ffcdd2');
            expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
        });

        it('saves the color chosen with a press, or with Space', async () => {
            const { worksheet, picker } = await create();

            picker.querySelector<HTMLElement>('td[data-value="#4caf50"]')!.click();
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValue('A1')).to.equal('#4caf50');

            worksheet.openEditor(worksheet.getCell('A1'));
            key(document.activeElement!, 'ArrowRight');
            key(document.activeElement!, ' ');
            expect(worksheet.getValue('A1')).to.equal('#8bc34a');
        });

        it('closes without saving on Escape', async () => {
            const { worksheet } = await create();

            key(document.activeElement!, 'ArrowRight');
            key(document.activeElement!, 'Escape');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValue('A1')).to.equal('#e57373');
        });

        it('removes the color with Reset, and saves the color of the input', async () => {
            const { worksheet, picker } = await create();

            picker.querySelector<HTMLButtonElement>('.jcolor-reset')!.click();
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValue('A1')).to.equal('');

            worksheet.openEditor(worksheet.getCell('A1'));
            const input = worksheet.getCell('A1').querySelector<HTMLInputElement>('.jcolor-custom input')!;
            input.value = '#123456';
            input.dispatchEvent(new Event('change'));
            expect(worksheet.getValue('A1')).to.equal('#123456');
        });

        it('saves when Done, a press outside or the focus leaving it closes it', async () => {
            const { worksheet, picker } = await create();
            picker.querySelector<HTMLButtonElement>('.jcolor-close')!.click();
            expect(worksheet.editor.isOpen()).to.equal(false);

            worksheet.openEditor(worksheet.getCell('A1'));
            document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(worksheet.editor.isOpen()).to.equal(false);

            worksheet.openEditor(worksheet.getCell('A1'));
            const outside = document.createElement('button');
            document.body.appendChild(outside);
            try {
                worksheet
                    .getCell('A1')
                    .querySelector('.jcolor')!
                    .dispatchEvent(new FocusEvent('focusout', { relatedTarget: outside }));
                expect(worksheet.editor.isOpen()).to.equal(false);
            } finally {
                outside.remove();
            }
            expect(worksheet.getValue('A1')).to.equal('#e57373');
        });

        it('takes the palette and the texts of the buttons from the options of the column', async () => {
            const { worksheet, picker } = await create('', { palette: [['ff0000', '#00ff00']], resetLabel: 'None', doneLabel: 'OK' });

            expect(Array.from(picker.querySelectorAll('td')).map((td) => td.getAttribute('data-value'))).to.eql(['#ff0000', '#00ff00']);
            expect(Array.from(picker.querySelectorAll('button')).map((button) => button.textContent)).to.eql(['None', 'OK']);
            // Without a color, the first color is the tab stop
            expect(document.activeElement).to.equal(picker.querySelector('td'));

            picker.querySelectorAll('td')[1].click();
            expect(worksheet.getValue('A1')).to.equal('#00ff00');
        });

        it('has a fixed position under the cell in a worksheet that scrolls', async () => {
            const instance = jspreadsheet(root, {
                worksheets: [{ data: [['#e57373']], columns: [{ type: 'color' }], tableOverflow: true }],
            });
            await waitForListeners();
            const worksheet = instance[0];
            worksheet.openEditor(worksheet.getCell('A1'));

            const content = worksheet.getCell('A1').querySelector<HTMLElement>('.jcolor-content')!;
            expect(content.style.position).to.equal('fixed');
            expect(content.style.top).to.equal('0px');
        });
    });

    describe('The colors of the toolbar', () => {
        const createWithToolbar = async () => {
            const instance = jspreadsheet(root, { toolbar: true, worksheets: [{ minDimensions: [3, 3] }] });
            await waitForListeners();
            instance[0].updateSelectionFromCoords(0, 0, 1, 0);
            return instance[0];
        };

        it('open with a press or a key, give the selected cells the color, and give the focus back', async () => {
            const worksheet = await createWithToolbar();
            const item = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Text color"]')!;

            item.focus();
            key(item, 'Enter');
            const picker = item.querySelector<HTMLElement>('.jcolor')!;
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            // The toolbar does not mark the color of the cells
            expect(picker.querySelector('[aria-selected="true"]')).to.equal(null);
            expect(document.activeElement).to.equal(picker.querySelector('td'));

            key(document.activeElement!, 'ArrowRight');
            key(document.activeElement!, 'Enter');
            expect(picker.classList.contains('jcolor-focus')).to.equal(false);
            expect(document.activeElement).to.equal(item);
            expect(worksheet.getCell('A1').style.color).to.equal('rgb(252, 228, 236)');
            expect(worksheet.getCell('B1').style.color).to.equal('rgb(252, 228, 236)');

            // It opens again, also with a key
            key(item, 'Enter');
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            key(document.activeElement!, 'Escape');
            expect(picker.classList.contains('jcolor-focus')).to.equal(false);
            expect(document.activeElement).to.equal(item);

            // A press on the item opens it, and a press on a color chooses it
            item.click();
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            item.querySelector<HTMLElement>('td[data-value="#4caf50"]')!.click();
            expect(picker.classList.contains('jcolor-focus')).to.equal(false);
            expect(worksheet.getCell('A1').style.color).to.equal('rgb(76, 175, 80)');
        });

        it('the color of the background colors the background', async () => {
            const worksheet = await createWithToolbar();
            const item = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Background color"]')!;

            item.click();
            item.querySelector<HTMLElement>('td[data-value="#2196f3"]')!.click();
            expect(worksheet.getCell('A1').style.backgroundColor).to.equal('rgb(33, 150, 243)');
        });

        it('the color of the borders stays open, and colors the borders', async () => {
            const worksheet = await createWithToolbar();
            const borders = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Borders"]')!;
            borders.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));

            const lens = borders.querySelector<HTMLElement>('[aria-label="Border color"]')!;
            lens.click();
            const picker = borders.querySelector<HTMLElement>('.jcolor')!;
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);

            picker.querySelector<HTMLElement>('td[data-value="#f44336"]')!.click();
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            expect(lens.style.color).to.equal('rgb(244, 67, 54)');
            picker.querySelector<HTMLButtonElement>('.jcolor-close')!.click();
            expect(picker.classList.contains('jcolor-focus')).to.equal(false);

            // The thickness and the style
            const [thickness, style] = Array.from(borders.querySelectorAll<HTMLElement>('.jpicker-content .jpicker'));
            expect(thickness.getAttribute('aria-label')).to.equal('Border thickness');
            expect(style.getAttribute('aria-label')).to.equal('Border style');
            thickness.querySelectorAll<HTMLElement>('.jpicker-item')[2].click();
            style.querySelectorAll<HTMLElement>('.jpicker-item')[1].click();

            // The borders picker stayed open: border_all
            expect(borders.classList.contains('jpicker-focus')).to.equal(true);
            borders.querySelectorAll<HTMLElement>(':scope > .jpicker-content > .jpicker-item')[0].click();
            expect(worksheet.getCell('A1').getAttribute('style')).to.contain('border-top: 3px dotted rgb(244, 67, 54)');

            // The key on the lens opens the color picker too
            lens.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
        });
    });
});
