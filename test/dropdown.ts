import { expect } from 'chai';

import gridwell, { type Gridwell, type Worksheet } from './support/gridwell';

// The key codes, which the keys of the editor and of the grid read
const codes: Record<string, number> = {
    Tab: 9,
    Enter: 13,
    Escape: 27,
    ' ': 32,
    PageUp: 33,
    PageDown: 34,
    End: 35,
    Home: 36,
    ArrowUp: 38,
    ArrowDown: 40,
};

const key = (element: Element, keyName: string) => {
    const keyCode = codes[keyName] ?? keyName.toUpperCase().charCodeAt(0);
    const event = new KeyboardEvent('keydown', { key: keyName, keyCode, which: keyCode, bubbles: true, cancelable: true });
    element.dispatchEvent(event);
    return event;
};

const press = (element: Element) => {
    element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
};

const waitForListeners = () => new Promise((resolve) => setTimeout(resolve));

const cars: Gridwell.DropdownSourceItem[] = [
    { id: 1, name: 'Mazda' },
    { id: 2, name: 'Honda' },
    { id: 3, name: 'Peugeot' },
    { id: 4, name: 'Hyundai' },
];

/** A worksheet with a dropdown column, with the editor of A1 open. */
async function open(column: Gridwell.DropdownColumn, value: Gridwell.CellValue = 2, options: Partial<Gridwell.WorksheetOptions> = {}) {
    const instance = gridwell(root, {
        worksheets: [
            {
                data: [
                    [value, 'x'],
                    ['', 'y'],
                ],
                columns: [{ type: 'dropdown', source: cars, ...column }, {}],
                ...options,
            },
        ],
    });
    await waitForListeners();
    const worksheet = instance[0];
    worksheet.updateSelectionFromCoords(0, 0, 0, 0);
    const cell = worksheet.getCell('A1');
    worksheet.openEditor(cell);
    return { worksheet, cell, ...parts(cell) };
}

/** The elements of a dropdown in an element. */
function parts(element: HTMLElement) {
    const header = element.querySelector<HTMLInputElement>('.jdropdown-header')!;
    const list = element.querySelector<HTMLElement>('.jdropdown-content')!;
    const items = () => Array.from(list.querySelectorAll<HTMLElement>('.jdropdown-item'));
    const active = () => document.getElementById(header.getAttribute('aria-activedescendant') || '');
    const shown = () => items().filter((item) => !item.hidden && !(item.parentElement!.parentElement as HTMLElement).hidden);
    return { header, list, items, active, shown };
}

const label = (item: Element | null) => item?.querySelector('.jdropdown-description')!.firstChild!.textContent;

describe('The dropdown', () => {
    describe('The editor of a dropdown column', () => {
        it('is a combobox with a listbox of options, and the chosen option is active', async () => {
            const { cell, header, list, items, active } = await open({});

            const dropdown = cell.querySelector('.jdropdown')!;
            expect(dropdown.classList.contains('jdropdown-focus')).to.equal(true);
            expect(document.activeElement).to.equal(header);
            expect(header.getAttribute('role')).to.equal('combobox');
            expect(header.getAttribute('aria-expanded')).to.equal('true');
            expect(header.getAttribute('aria-haspopup')).to.equal('listbox');
            expect(header.getAttribute('aria-controls')).to.equal(list.id);
            expect(header.getAttribute('aria-label')).to.equal('Edit cell A1');
            expect(header.readOnly).to.equal(true);
            expect(header.hasAttribute('aria-autocomplete')).to.equal(false);
            expect(header.value).to.equal('Honda');

            expect(list.getAttribute('role')).to.equal('listbox');
            expect(list.hasAttribute('aria-multiselectable')).to.equal(false);
            expect(items().map((item) => item.getAttribute('role'))).to.eql(['option', 'option', 'option', 'option']);
            expect(items().map((item) => item.textContent)).to.eql(['Mazda', 'Honda', 'Peugeot', 'Hyundai']);
            expect(items().map((item) => item.getAttribute('aria-selected'))).to.eql(['false', 'true', 'false', 'false']);
            expect(items()[1].classList.contains('jdropdown-selected')).to.equal(true);

            expect(active()).to.equal(items()[1]);
            expect(items()[1].classList.contains('jdropdown-cursor')).to.equal(true);
        });

        it('moves through the options with the arrows, Page Up, Page Down, Home, End and a letter', async () => {
            const { header, active } = await open({});

            key(header, 'ArrowDown');
            expect(label(active())).to.equal('Peugeot');
            key(header, 'ArrowUp');
            key(header, 'ArrowUp');
            expect(label(active())).to.equal('Mazda');
            // The first and the last option stop the arrows
            key(header, 'ArrowUp');
            expect(label(active())).to.equal('Mazda');
            key(header, 'PageDown');
            expect(label(active())).to.equal('Hyundai');
            key(header, 'PageUp');
            expect(label(active())).to.equal('Mazda');
            key(header, 'End');
            expect(label(active())).to.equal('Hyundai');
            expect(key(header, 'Home').defaultPrevented).to.equal(true);
            expect(label(active())).to.equal('Mazda');

            // A letter goes to the next option that starts with it, and then the one after
            key(header, 'h');
            expect(label(active())).to.equal('Honda');
            key(header, 'h');
            expect(label(active())).to.equal('Hyundai');
            key(header, 'h');
            expect(label(active())).to.equal('Honda');
            key(header, 'z');
            expect(label(active())).to.equal('Honda');
        });

        it('saves the option chosen with Enter, and keeps the selection on the cell', async () => {
            const { worksheet, header } = await open({});
            const changes: unknown[] = [];
            worksheet.parent.config.onchange = (instance, cell, x, y, value) => void changes.push(value);

            key(header, 'ArrowDown');
            const event = key(header, 'Enter');

            expect(event.defaultPrevented).to.equal(true);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('3');
            expect(worksheet.getCell('A1').textContent).to.equal('Peugeot');
            expect(changes).to.eql(['3']);
            expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
        });

        it('keeps the chosen option with Enter, and keeps the value with Enter when no option is active', async () => {
            const { worksheet, header } = await open({});

            key(header, 'Enter');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal(2);

            // A cell without a value has no active option
            worksheet.updateSelectionFromCoords(0, 1, 0, 1);
            worksheet.openEditor(worksheet.getCell('A2'));
            const second = parts(worksheet.getCell('A2'));
            expect(second.active()).to.equal(null);
            expect(second.header.value).to.equal('');
            key(second.header, 'Enter');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 1)).to.equal('');
        });

        it('saves the option that is pressed, and a press on the chosen option empties the cell', async () => {
            const { worksheet, items } = await open({});

            press(items()[0]);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('1');

            worksheet.openEditor(worksheet.getCell('A1'));
            press(parts(worksheet.getCell('A1')).items()[0].querySelector('.jdropdown-description')!);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('');
        });

        it('closes without saving with Escape', async () => {
            const { worksheet, header } = await open({});
            const ends: unknown[] = [];
            worksheet.parent.config.oneditionend = (instance, cell, x, y, value, save) => void ends.push(save);

            key(header, 'ArrowDown');
            key(header, 'Escape');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal(2);
            expect(worksheet.getCell('A1').textContent).to.equal('Honda');
            expect(ends).to.eql([false]);
        });

        it('saves the active option with Tab, and moves the selection to the next cell', async () => {
            const { worksheet, header } = await open({});

            key(header, 'ArrowDown');
            key(header, 'Tab');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('3');
            expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);
        });

        it('saves the value after a press outside it, or a press on its header', async () => {
            const { worksheet, header } = await open({});

            key(header, 'ArrowDown');
            press(header);
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal(2);

            worksheet.openEditor(worksheet.getCell('A1'));
            const outside = document.createElement('button');
            document.body.appendChild(outside);
            outside.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            outside.remove();

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal(2);
        });

        it('adds and removes options with Space, Enter and a press when it has multiple, and saves them with Tab', async () => {
            const { worksheet, header, list, items, active } = await open({ multiple: true, delimiter: ', ' }, '2, 9');

            expect(list.getAttribute('aria-multiselectable')).to.equal('true');
            expect(items().map((item) => item.getAttribute('aria-selected'))).to.eql(['false', 'true', 'false', 'false']);
            // A value without an option stays
            expect(header.value).to.equal('Honda; 9');

            key(header, 'ArrowUp');
            expect(label(active())).to.equal('Mazda');
            expect(key(header, ' ').defaultPrevented).to.equal(true);
            key(header, 'ArrowDown');
            key(header, 'Enter');
            press(items()[3]);

            expect(worksheet.editor.isOpen()).to.equal(true);
            expect(items().map((item) => item.getAttribute('aria-selected'))).to.eql(['true', 'false', 'false', 'true']);
            expect(header.value).to.equal('9; Mazda; Hyundai');

            key(header, 'Tab');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('9, 1, 4');
        });

        it('filters the options with the text typed in it when it has autocomplete', async () => {
            const { worksheet, header, list, items, shown, active } = await open({ autocomplete: true });

            expect(header.readOnly).to.equal(false);
            expect(header.getAttribute('aria-autocomplete')).to.equal('list');
            // The header has the text that filters the options
            expect(header.value).to.equal('');

            header.value = 'HY';
            header.dispatchEvent(new Event('input'));
            // The chosen option stays in the list
            expect(shown().map(label)).to.eql(['Honda', 'Hyundai']);
            expect(active()).to.equal(null);

            // Home and End move in the text, and Space types a space
            expect(key(header, 'End').defaultPrevented).to.equal(false);
            expect(label(active())).to.equal('Hyundai');
            expect(key(header, 'Home').defaultPrevented).to.equal(false);
            expect(label(active())).to.equal('Honda');
            expect(key(header, ' ').defaultPrevented).to.equal(false);
            expect(key(header, 'q').defaultPrevented).to.equal(false);

            header.value = 'zzz';
            header.dispatchEvent(new Event('input'));
            expect(shown()).to.eql([items()[1]]);

            expect(list.style.display).to.equal('');

            header.value = 'yun';
            header.dispatchEvent(new Event('input'));
            key(header, 'ArrowDown');
            key(header, 'ArrowDown');
            key(header, 'Enter');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal('4');
        });

        it('hides the list when no option has the text that is typed', async () => {
            const { header, list } = await open({ autocomplete: true }, '');

            header.value = 'xyz';
            header.dispatchEvent(new Event('input'));

            expect(list.style.display).to.equal('none');
            key(header, 'ArrowDown');
            expect(header.hasAttribute('aria-activedescendant')).to.equal(false);
        });

        it('shows the groups, the images and the titles of the options, as text', async () => {
            const source: Gridwell.DropdownSourceItem[] = [
                'Plain',
                7,
                { id: 'a', name: '<b>Bold</b>', title: 'A title', group: 'Group 1' },
                { id: 'b', name: '', image: 'data:image/png;base64,', group: 'Group 1' },
                { id: 'c', name: 'Image', image: 'data:image/png;base64,', title: 'With title' },
            ];
            const { list, items, header } = await open({ source, autocomplete: true }, 'a');

            const group = list.querySelector('.jdropdown-group')!;
            expect(group.getAttribute('role')).to.equal('group');
            expect(document.getElementById(group.getAttribute('aria-labelledby')!)!.textContent).to.equal('Group 1');
            // The options of the groups come first
            expect(items().map(label)).to.eql(['<b>Bold</b>', ' ', 'Plain', '7', 'Image']);
            expect(list.querySelector('b')).to.equal(null);
            expect(items()[0].querySelector('.jdropdown-title')!.textContent).to.equal('A title');
            const images = list.querySelectorAll('img');
            expect(images[0].classList.contains('jdropdown-image-small')).to.equal(true);
            expect(images[1].classList.contains('jdropdown-image-small')).to.equal(false);
            expect(images[0].alt).to.equal('');

            // A search finds the titles and the groups, and hides a group without options
            header.value = 'with';
            header.dispatchEvent(new Event('input'));
            expect((group as HTMLElement).hidden).to.equal(false);
            header.value = 'Group';
            header.dispatchEvent(new Event('input'));
            expect(
                items()
                    .filter((item) => !item.hidden)
                    .map(label)
            ).to.eql(['<b>Bold</b>', ' ']);
            header.value = '7';
            header.dispatchEvent(new Event('input'));
            expect(
                items()
                    .filter((item) => !item.hidden)
                    .map(label)
            ).to.eql(['<b>Bold</b>', '7']);
        });

        it('opens its list with the arrows and with a press on its header when it is closed', async () => {
            const { worksheet, cell, header } = await open({ autocomplete: true });
            const dropdown = cell.querySelector('.jdropdown') as HTMLElement & { dropdown: { close(ignoreEvents: boolean): string } };

            expect(dropdown.dropdown.close(true)).to.equal('2');
            expect(header.getAttribute('aria-expanded')).to.equal('false');
            expect(header.value).to.equal('Honda');
            expect(key(header, 'Enter').defaultPrevented).to.equal(false);
            expect(key(header, 'q').defaultPrevented).to.equal(false);

            key(header, 'ArrowDown');
            expect(header.getAttribute('aria-expanded')).to.equal('true');
            dropdown.dropdown.close(true);
            press(header);
            expect(header.getAttribute('aria-expanded')).to.equal('true');

            // With autocomplete, a press on the header closes it only on the arrow, at its right
            header.getBoundingClientRect = () => ({ left: 0, width: 200 }) as DOMRect;
            header.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 10 }));
            expect(header.getAttribute('aria-expanded')).to.equal('true');
            header.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 190 }));
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(0, 0)).to.equal(2);
        });

        it('gives the list a fixed position in a worksheet that scrolls', async () => {
            const { cell } = await open({}, 2, { tableOverflow: true });

            expect(cell.querySelector<HTMLElement>('.jdropdown-container')!.style.position).to.equal('fixed');
        });
    });

    describe('The place of the list', () => {
        const original = HTMLElement.prototype.getBoundingClientRect;

        afterEach(() => {
            HTMLElement.prototype.getBoundingClientRect = original;
            root.removeAttribute('style');
        });

        /**
         * Give the dropdown, its list and the root the sizes of a layout, which jsdom does not have. The window is 768 pixels
         * high.
         */
        const layout = (cell: [number, number], list: number, container?: [number, number]) => {
            HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
                if (this.classList.contains('jdropdown')) {
                    return { top: cell[0], bottom: cell[1], height: cell[1] - cell[0], left: 10, width: 100 } as DOMRect;
                }
                if (this.classList.contains('jdropdown-container')) {
                    return { height: list } as DOMRect;
                }
                if (this === root && container) {
                    return { top: container[0], bottom: container[1] } as DOMRect;
                }
                return original.call(this);
            };
        };

        it('opens the list under the cell when it fits', async () => {
            layout([100, 126], 215);
            const { list } = await open({});

            expect(list.parentElement!.style.bottom).to.equal('');
            expect(list.style.maxHeight).to.equal('');
        });

        it('opens the list over the cell when there is no room under it in the window', async () => {
            layout([700, 726], 215);
            const { list } = await open({});

            expect(list.parentElement!.style.bottom).to.equal('27px');
        });

        it('opens the list over the cell when the element that scrolls has no room under it', async () => {
            root.style.cssText = 'position: relative; overflow: auto';
            layout([250, 276], 215, [0, 300]);
            const { list } = await open({});

            expect(list.parentElement!.style.bottom).to.equal('27px');
        });

        it('opens the list under the cell when it fits on neither side and there is more room under it, and the list scrolls', async () => {
            layout([300, 326], 600);
            const { list } = await open({});

            expect(list.parentElement!.style.bottom).to.equal('');
            expect(list.style.maxHeight).to.equal('440px');
        });

        it('opens the list over the cell when it fits on neither side and there is more room over it, and the list scrolls', async () => {
            layout([450, 476], 600);
            const { list } = await open({});

            expect(list.parentElement!.style.bottom).to.equal('27px');
            expect(list.style.maxHeight).to.equal('448px');
        });

        it('places a list with a fixed position under the cell', async () => {
            layout([100, 126], 215);
            const { list } = await open({}, 2, { tableOverflow: true });

            expect(list.parentElement!.style.top).to.equal('126px');
            expect(list.parentElement!.style.left).to.equal('10px');
        });

        it('places a list with a fixed position over the cell when there is no room under it in the window', async () => {
            layout([700, 726], 215);
            const { list } = await open({}, 2, { tableOverflow: true });

            expect(list.parentElement!.style.top).to.equal('');
            expect(list.parentElement!.style.bottom).to.equal('69px');
        });
    });

    describe('The filters of the columns', () => {
        const data = [
            ['Honda', 1],
            ['Mazda', 2],
            ['Honda', 3],
            ['', 4],
        ];

        const openFilter = async (worksheet: Worksheet, column = 0) => {
            worksheet.openFilter(column);
            return parts(worksheet.filter!.children[column + 1] as HTMLElement);
        };

        it('lists the values of the column, has a name and filters the rows with the chosen values', async () => {
            const worksheet = gridwell(root, { worksheets: [{ filters: true, data, columns: [{ title: 'Car' }, {}] }] })[0];
            await waitForListeners();
            const { header, items, list } = await openFilter(worksheet);

            expect(header.getAttribute('aria-label')).to.equal('Filter column Car');
            expect(list.getAttribute('aria-multiselectable')).to.equal('true');
            expect(items().map(label)).to.eql(['Honda', 'Mazda', '(Blanks)']);

            press(items()[0]);
            key(header, 'ArrowDown');
            key(header, 'ArrowDown');
            key(header, 'Enter');
            expect(items().map((item) => item.getAttribute('aria-selected'))).to.eql(['true', 'false', 'true']);

            key(header, 'Tab');

            expect(worksheet.view.filters[0]).to.eql(['Honda', '']);
            expect(worksheet.view.results).to.eql([0, 2, 3]);
            expect(worksheet.filter!.children[1].textContent).to.equal('Honda; (Blanks)');
            expect(worksheet.filter!.children[1].querySelector('.jdropdown')).to.equal(null);
        });

        it('keeps the filters as they were after Escape', async () => {
            const worksheet = gridwell(root, { worksheets: [{ filters: true, data }] })[0];
            await waitForListeners();
            let filter = await openFilter(worksheet);
            press(filter.items()[1]);
            key(filter.header, 'Tab');
            expect(worksheet.view.results).to.eql([1]);

            filter = await openFilter(worksheet);
            expect(filter.items().map((item) => item.getAttribute('aria-selected'))).to.eql(['false', 'true', 'false']);
            press(filter.items()[0]);
            key(filter.header, 'Escape');

            expect(worksheet.view.filters[0]).to.eql(['Mazda']);
            expect(worksheet.view.results).to.eql([1]);
            const cell = worksheet.filter!.children[1] as HTMLElement;
            expect(cell.textContent).to.equal('Mazda');
            expect(cell.style.overflow).to.equal('');
        });
    });
});
