import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

describe('Toolbar', () => {
    const createWithToolbar = async () => {
        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [{ minDimensions: [3, 3] }],
        });

        // The spreadsheet creates the toolbar after its worksheets
        await new Promise((resolve) => setTimeout(resolve, 0));

        return instance;
    };

    describe('Font', () => {
        it('Default removes the font of the selected cells', async () => {
            const instance = await createWithToolbar();

            instance[0].updateSelectionFromCoords(0, 0, 0, 0);

            const fonts = root.querySelector('.jss_toolbar .jpicker')!.querySelectorAll<HTMLElement>('.jpicker-item');
            expect(fonts[0].textContent).to.equal('Default');
            expect(fonts[1].textContent).to.equal('Verdana');

            fonts[1].click();
            expect(instance[0].getCell('A1').style.fontFamily.trim()).to.equal('Verdana');

            fonts[0].click();
            expect(instance[0].getCell('A1').style.fontFamily.trim()).to.equal('');
        });
    });

    describe('Default items', () => {
        // The icon item whose icon has the text
        const getIconItem = (icon: string) =>
            Array.from(root.querySelectorAll<HTMLElement>('.jss_toolbar .jtoolbar-item')).find((item) => item.children[0]?.textContent === icon)!;

        // Choose the option at index in the picker at position among the select items
        const choose = (position: number, index: number) => {
            const picker = root.querySelectorAll('.jss_toolbar .jpicker')[position];
            picker.querySelectorAll<HTMLElement>('.jpicker-item')[index].click();
        };

        const getStyle = (worksheet: Jspreadsheet.WorksheetInstance, cellName: string) => worksheet.getCell(cellName).getAttribute('style') || '';

        it('has the items of the default toolbar, in order', async () => {
            let items = null as unknown[] | null;
            jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3] }],
                toolbar(toolbar) {
                    // A divisor item has no content
                    items = toolbar.items.map((item) => (item as { content?: string }).content || item.type);
                    return toolbar;
                },
            });
            await new Promise((resolve) => setTimeout(resolve, 0));

            expect(items).to.eql([
                'undo',
                'redo',
                'save',
                'divisor',
                'select',
                'format_size',
                'select',
                'format_bold',
                'format_color_text',
                'format_color_fill',
                'select',
                'web',
                'border_all',
                'divisor',
                'fullscreen',
            ]);
        });

        it('the size, the alignment, bold and the vertical alignment style the selected cells', async () => {
            const changes: Record<string, string>[] = [];
            jspreadsheet(root, {
                toolbar: true,
                worksheets: [{ minDimensions: [3, 3] }],
                onchangestyle: (worksheet, newValue) => {
                    changes.push(newValue);
                },
            });
            await new Promise((resolve) => setTimeout(resolve, 0));
            const worksheet = jspreadsheet.current!;
            worksheet.updateSelectionFromCoords(0, 0, 1, 1);

            choose(1, 3);
            // The cells are centered already: choosing center keeps them centered, and changes nothing
            choose(2, 1);
            getIconItem('format_bold').click();
            choose(3, 0);

            // The toolbar gives the values with a space after the colon, and setStyle trims them (issue #67)
            expect(changes.map((change) => Object.keys(change).sort())).to.eql([
                ['A1', 'A2', 'B1', 'B2'],
                ['A1', 'A2', 'B1', 'B2'],
                ['A1', 'A2', 'B1', 'B2'],
            ]);
            ['A1', 'B1', 'A2', 'B2'].forEach((cellName) => {
                expect(changes[0][cellName], cellName).to.equal('font-size:large');
                expect(changes[1][cellName], cellName).to.equal('font-weight:bold');
                expect(changes[2][cellName], cellName).to.equal('vertical-align:top');

                const style = getStyle(worksheet, cellName);
                expect(style, cellName).to.contain('font-size: large');
                expect(style, cellName).to.contain('text-align: center');
                expect(style, cellName).to.contain('font-weight: bold');
                expect(style, cellName).to.contain('vertical-align: top');
            });

            // Bold toggles off
            getIconItem('format_bold').click();
            expect(getStyle(worksheet, 'A1')).not.to.contain('font-weight');

            // The other cells keep the alignment of their column
            expect(getStyle(worksheet, 'C3')).to.equal('text-align: center;');
        });

        it('merge merges the selected cells, and removes the merge of merged cells', async () => {
            // The toolbar calls the global confirm, which is not the confirm of window in the jsdom of the tests
            const confirm = globalThis.confirm;
            globalThis.confirm = () => true;
            try {
                const instance = await createWithToolbar();
                const worksheet = instance[0];

                worksheet.updateSelectionFromCoords(1, 1, 0, 0);
                getIconItem('web').click();
                expect(worksheet.getMerge('A1')).to.eql([2, 2]);

                getIconItem('web').click();
                expect(worksheet.getMerge('A1')).to.eql(null);
            } finally {
                globalThis.confirm = confirm;
            }
        });

        it('the borders set the borders of the selected cells', async () => {
            const instance = await createWithToolbar();
            const worksheet = instance[0];
            worksheet.updateSelectionFromCoords(0, 0, 1, 1);

            const borders = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Borders"]')!;
            const chooseBorders = (label: string) => {
                borders.click();
                document.getElementById(borders.getAttribute('aria-controls')!)!.querySelector<HTMLElement>(`button[aria-label="${label}"]`)!.click();
            };

            chooseBorders('Outer borders');
            expect(getStyle(worksheet, 'A1')).to.contain('border-top: 1px solid black').and.to.contain('border-left: 1px solid black');
            expect(getStyle(worksheet, 'A1')).not.to.contain('border-right: 1px').and.not.to.contain('border-bottom: 1px');
            expect(getStyle(worksheet, 'B2')).to.contain('border-right: 1px solid black').and.to.contain('border-bottom: 1px solid black');
            expect(getStyle(worksheet, 'B2')).not.to.contain('border-top: 1px').and.not.to.contain('border-left: 1px');

            chooseBorders('Inner borders');
            expect(getStyle(worksheet, 'B2')).to.contain('border-top: 1px solid black').and.to.contain('border-left: 1px solid black');
            expect(getStyle(worksheet, 'A1')).not.to.contain('border-top: 1px').and.not.to.contain('border-left: 1px');
        });

        it('fullscreen toggles the fullscreen of the spreadsheet and its icon', async () => {
            const instance = await createWithToolbar();
            const item = getIconItem('fullscreen');

            item.click();
            expect(instance[0].parent.config.fullscreen).to.equal(true);
            expect(item.children[0].textContent).to.equal('fullscreen_exit');

            item.click();
            expect(instance[0].parent.config.fullscreen).to.equal(false);
            expect(item.children[0].textContent).to.equal('fullscreen');
        });

        it('the items that change cells are disabled on a worksheet that is not editable', async () => {
            jspreadsheet(root, {
                toolbar: true,
                worksheets: [{ minDimensions: [3, 3], editable: false }],
            });
            await new Promise((resolve) => setTimeout(resolve, 0));

            expect(getIconItem('format_bold').classList.contains('jtoolbar-disabled')).to.equal(true);
            expect(getIconItem('web').classList.contains('jtoolbar-disabled')).to.equal(true);
            expect(getIconItem('undo').classList.contains('jtoolbar-disabled')).to.equal(false);
        });
    });

    describe('Settings', () => {
        it('gives a toolbar function and the toolbar of a plugin the settings with the items', async () => {
            const calls: [string, boolean | number][] = [];
            jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3] }],
                toolbar(toolbar) {
                    calls.push(['option', Array.isArray(toolbar.items)]);
                    toolbar.items = toolbar.items.filter(
                        (item) => (item as { content?: string }).content === 'undo' || (item as { content?: string }).content === 'redo'
                    );
                    return toolbar;
                },
                plugins: {
                    test() {
                        return {
                            toolbar(toolbar) {
                                calls.push(['plugin', toolbar.items.length]);
                            },
                        };
                    },
                },
            });

            await new Promise((resolve) => setTimeout(resolve, 0));

            expect(calls).to.eql([
                ['option', true],
                ['plugin', 2],
            ]);
            expect(root.querySelectorAll('.jss_toolbar .jtoolbar-item').length).to.equal(2);
        });
    });

    describe('Show, hide and update', () => {
        it('hideToolbar removes the toolbar, and showToolbar adds it again', async () => {
            const instance = await createWithToolbar();
            const spreadsheet = instance[0].parent;

            expect(spreadsheet.toolbar).to.equal(root.querySelector('.jss_toolbar'));

            spreadsheet.hideToolbar();

            expect(root.querySelector('.jss_toolbar')).to.equal(null);
            expect(spreadsheet.toolbar).to.equal(undefined);

            spreadsheet.showToolbar();

            expect(root.querySelectorAll('.jss_toolbar').length).to.equal(1);
            expect(spreadsheet.toolbar).to.equal(root.querySelector('.jss_toolbar'));
            expect(spreadsheet.toolbar.querySelectorAll('.jtoolbar-item').length).to.equal(15);
        });

        it('setConfig with toolbar hides and shows the toolbar', async () => {
            const instance = await createWithToolbar();
            const spreadsheet = instance[0].parent;

            spreadsheet.setConfig({ toolbar: false });
            expect(root.querySelector('.jss_toolbar')).to.equal(null);

            spreadsheet.setConfig({ toolbar: true });
            expect(root.querySelector('.jss_toolbar')).to.equal(spreadsheet.toolbar);
        });

        it('gives updateState the worksheet whose cells are selected, and the worksheet of the open tab', async () => {
            const states: Jspreadsheet.WorksheetInstance[] = [];
            const instance = jspreadsheet(root, {
                toolbar: [
                    {
                        type: 'i',
                        content: 'info',
                        tooltip: 'Information',
                        updateState: (toolbarElement, toolbarInstance, itemElement, worksheet) => {
                            states.push(worksheet);
                        },
                    },
                ],
                worksheets: [{ minDimensions: [3, 3] }, { minDimensions: [3, 3] }],
            });
            await new Promise((resolve) => setTimeout(resolve, 0));

            // The tooltip is the title of the item
            expect(root.querySelector<HTMLElement>('.jss_toolbar .jtoolbar-item')!.title).to.equal('Information');

            states.length = 0;
            instance[0].updateSelectionFromCoords(0, 0, 0, 0);
            expect(states[states.length - 1]).to.equal(instance[0]);

            states.length = 0;
            (root as Worksheet['parent']['element']).tabs!.open(1);
            expect(states).to.eql([instance[1]]);
        });

        it('a select item with a style property styles the selected cells with the chosen value', async () => {
            const changes: Record<string, string>[] = [];
            const instance = jspreadsheet(root, {
                toolbar: [{ type: 'select', k: 'text-align', v: ['left', 'right'] }],
                worksheets: [{ minDimensions: [3, 3] }],
                onchangestyle: (worksheet, newValue) => {
                    changes.push(newValue);
                },
            });
            await new Promise((resolve) => setTimeout(resolve, 0));

            instance[0].updateSelectionFromCoords(0, 0, 1, 0);
            root.querySelector('.jss_toolbar .jpicker')!.querySelectorAll<HTMLElement>('.jpicker-item')[1].click();

            expect(changes).to.eql([{ A1: 'text-align:right', B1: 'text-align:right' }]);
        });
    });

    describe('Items', () => {
        it('creates labels, icons with other classes, and items with an id, a width, a title, and active or disabled', async () => {
            const clicks: unknown[][] = [];
            const instance = jspreadsheet(root, {
                toolbar: [
                    { type: 'label', content: '<b>Label</b>' },
                    { type: 'divisor' },
                    {
                        content: 'star',
                        class: 'fa fa-star',
                        id: 'star',
                        width: '40px',
                        title: 'Star',
                        active: true,
                        render(element: HTMLElement) {
                            element.dataset.rendered = 'yes';
                        },
                        onclick(...args: unknown[]) {
                            clicks.push(args);
                        },
                    },
                    { content: 'block', tooltip: 'Disabled', disabled: true },
                    { content: 'info' },
                ],
                worksheets: [{ minDimensions: [3, 3] }],
            });
            await new Promise((resolve) => setTimeout(resolve, 0));

            const items = root.querySelectorAll<HTMLElement>('.jss_toolbar .jtoolbar-item');
            // The content of a label is text
            expect(items[0].classList.contains('jtoolbar-label')).to.equal(true);
            expect(items[0].textContent).to.equal('<b>Label</b>');
            expect(items[1].getAttribute('role')).to.equal('separator');

            const star = items[2];
            expect(star.id).to.equal('star');
            expect(star.style.width).to.equal('40px');
            expect(star.title).to.equal('Star');
            expect(star.dataset.rendered).to.equal('yes');
            expect(star.classList.contains('jtoolbar-active')).to.equal(true);
            expect(star.children[0].className).to.equal('fa fa-star');
            // An icon without a tooltip is the name of its item
            expect(star.children[0].hasAttribute('aria-hidden')).to.equal(false);

            expect(items[3].classList.contains('jtoolbar-disabled')).to.equal(true);
            expect(items[3].getAttribute('aria-disabled')).to.equal('true');
            expect(items[3].children[0].getAttribute('aria-hidden')).to.equal('true');

            // onclick gets the toolbar, its instance, the item and the click, and the item is selected
            star.click();
            const toolbar = instance[0].parent.toolbar;
            expect(clicks[0].slice(0, 3)).to.eql([toolbar, toolbar.toolbar, star]);
            expect(clicks[0][3]).to.be.instanceOf(MouseEvent);
            expect(star.classList.contains('jtoolbar-selected')).to.equal(true);
            items[4].click();
            expect(star.classList.contains('jtoolbar-selected')).to.equal(false);
        });

        it('has the methods of the toolbar on the element', async () => {
            const instance = await createWithToolbar();
            const element = instance[0].parent.toolbar;
            const toolbar = element.toolbar as unknown as Jspreadsheet.ToolbarInstance;

            expect(toolbar.get()).to.equal(element);
            expect((toolbar.options.items as unknown[]).length).to.equal(15);

            toolbar.hide();
            expect(element.style.display).to.equal('none');
            toolbar.show();
            expect(element.style.display).to.equal('');

            toolbar.setReadonly(true);
            expect(element.classList.contains('jtoolbar-disabled')).to.equal(true);
            toolbar.setReadonly(false);
            expect(element.classList.contains('jtoolbar-disabled')).to.equal(false);

            toolbar.create([{ content: 'undo', tooltip: 'Undo' }, { content: 'redo' }]);
            const items = element.querySelectorAll<HTMLElement>('.jtoolbar-item');
            expect(items.length).to.equal(2);
            expect(Array.from(items).map((item) => item.getAttribute('tabindex'))).to.eql(['0', '-1']);

            toolbar.destroy();
            expect(element.querySelectorAll('.jtoolbar-item').length).to.equal(0);
        });

        it('marks the chosen option of a picker', async () => {
            await createWithToolbar();

            const font = root.querySelector<HTMLElement>('.jss_toolbar .jpicker')!;
            const fonts = font.querySelectorAll<HTMLElement>('.jpicker-item');
            expect(fonts[0].getAttribute('aria-selected')).to.equal('true');
            expect(font.querySelector('.jpicker-header')!.textContent).to.equal('Default');

            fonts[2].click();
            expect(fonts[0].getAttribute('aria-selected')).to.equal('false');
            expect(fonts[2].getAttribute('aria-selected')).to.equal('true');
            expect(font.querySelector('.jpicker-header')!.textContent).to.equal('Arial');
            expect((font as HTMLElement & { picker: Jspreadsheet.PickerInstance }).picker.getValue()).to.equal('2');
        });

        it('closes a picker on a press outside it', async () => {
            await createWithToolbar();

            const picker = root.querySelector<HTMLElement>('.jss_toolbar .jpicker')!;
            const instance = (picker as HTMLElement & { picker: Jspreadsheet.PickerInstance }).picker;
            picker.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(instance.isOpened()).to.equal(true);

            // A press in the picker keeps it open
            picker.querySelector('.jpicker-content')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(instance.isOpened()).to.equal(true);

            document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(instance.isOpened()).to.equal(false);
            expect(picker.getAttribute('aria-expanded')).to.equal('false');
        });
    });

    describe('Keys', () => {
        const key = (element: Element, keyName: string, keyCode: number) => {
            const event = new KeyboardEvent('keydown', { key: keyName, keyCode, which: keyCode, bubbles: true, cancelable: true });
            element.dispatchEvent(event);
            return event;
        };

        it('keeps the keys on a focused picker away from the grid (issue #79)', async () => {
            const instance = await createWithToolbar();
            const worksheet = instance[0];
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);

            const picker = root.querySelector<HTMLElement>('.jss_toolbar .jpicker')!;
            picker.focus();

            key(picker, 'a', 65);
            expect(worksheet.editor.isOpen()).to.equal(false);

            key(picker, 'Delete', 46);
            key(picker, 'ArrowRight', 39);
            expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
        });

        it('opens a picker and chooses an option with the keyboard', async () => {
            const instance = await createWithToolbar();
            instance[0].updateSelectionFromCoords(0, 0, 0, 0);

            const picker = root.querySelector<HTMLElement>('.jss_toolbar .jpicker')!;
            const fonts = picker.querySelectorAll('.jpicker-item');
            picker.focus();

            expect(key(picker, 'Enter', 13).defaultPrevented).to.equal(true);
            expect(picker.classList.contains('jpicker-focus')).to.equal(true);
            expect(picker.getAttribute('aria-expanded')).to.equal('true');
            expect(picker.getAttribute('aria-activedescendant')).to.equal(fonts[0].id);

            key(picker, 'ArrowDown', 40);
            expect(picker.getAttribute('aria-activedescendant')).to.equal(fonts[1].id);
            expect(fonts[1].classList.contains('jss_picker_active')).to.equal(true);

            key(picker, 'Enter', 13);
            expect(instance[0].getCell('A1').style.fontFamily.trim()).to.equal('Verdana');
            expect(picker.classList.contains('jpicker-focus')).to.equal(false);
            expect(picker.getAttribute('aria-expanded')).to.equal('false');
            expect(picker.hasAttribute('aria-activedescendant')).to.equal(false);

            // It opens on the chosen option
            key(picker, ' ', 32);
            expect(picker.getAttribute('aria-activedescendant')).to.equal(fonts[1].id);

            key(picker, 'End', 35);
            expect(picker.getAttribute('aria-activedescendant')).to.equal(fonts[fonts.length - 1].id);

            key(picker, 'Escape', 27);
            expect(picker.classList.contains('jpicker-focus')).to.equal(false);
            expect(instance[0].getCell('A1').style.fontFamily.trim()).to.equal('Verdana');
        });

        it('is one tab stop, and the arrows, Home and End move between its items (issue #125)', async () => {
            await createWithToolbar();

            const toolbar = root.querySelector('.jss_toolbar')!;
            expect(toolbar.getAttribute('role')).to.equal('toolbar');
            expect(toolbar.getAttribute('aria-label')).to.equal('Toolbar');

            const items = Array.from(toolbar.querySelectorAll<HTMLElement>('.jtoolbar-item[role="button"], .jtoolbar-item.jpicker'));
            // Undo, redo, save, the font, the size, the alignment, bold, the two colors, the vertical alignment, merge, the borders and fullscreen
            expect(items.length).to.equal(13);
            expect(items.map((item) => item.getAttribute('tabindex'))).to.eql(['0', ...Array(12).fill('-1')]);

            items[0].focus();
            expect(key(items[0], 'ArrowRight', 39).defaultPrevented).to.equal(true);
            expect(document.activeElement).to.equal(items[1]);
            expect(items[0].getAttribute('tabindex')).to.equal('-1');
            expect(items[1].getAttribute('tabindex')).to.equal('0');

            key(items[1], 'End', 35);
            expect(document.activeElement).to.equal(items[12]);
            key(items[12], 'ArrowRight', 39);
            expect(document.activeElement).to.equal(items[0]);
            key(items[0], 'ArrowLeft', 37);
            expect(document.activeElement).to.equal(items[12]);
            key(items[12], 'Home', 36);
            expect(document.activeElement).to.equal(items[0]);

            // The arrows move past a closed picker, and stay in an open one
            const picker = toolbar.querySelector<HTMLElement>('.jpicker')!;
            picker.focus();
            key(picker, 'ArrowRight', 39);
            expect(document.activeElement).to.not.equal(picker);
            picker.focus();
            key(picker, 'Enter', 13);
            key(picker, 'ArrowRight', 39);
            expect(document.activeElement).to.equal(picker);
            key(picker, 'Escape', 27);
        });

        it('opens the borders in a dialog, where the keys reach the types, the color, the thickness and the style (issue #146)', async () => {
            const instance = await createWithToolbar();
            const worksheet = instance[0];
            worksheet.updateSelectionFromCoords(0, 0, 1, 1);

            const borders = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Borders"]')!;
            expect(borders.getAttribute('role')).to.equal('button');
            expect(borders.getAttribute('aria-haspopup')).to.equal('dialog');
            expect(borders.getAttribute('aria-expanded')).to.equal('false');

            // Enter opens the dialog, and moves the focus to the first type of borders
            borders.focus();
            key(borders, 'Enter', 13);
            expect(borders.getAttribute('aria-expanded')).to.equal('true');
            const dialog = document.getElementById(borders.getAttribute('aria-controls')!)!;
            expect(dialog.getAttribute('role')).to.equal('dialog');
            expect(dialog.getAttribute('aria-label')).to.equal('Borders');
            expect(dialog.hidden).to.equal(false);
            // The dialog is not in the button, whose content is not read
            expect(borders.contains(dialog)).to.equal(false);

            const types = Array.from(dialog.querySelectorAll<HTMLElement>('[role="group"] > button'));
            expect(types.map((type) => type.getAttribute('aria-label'))).to.eql([
                'All borders',
                'Outer borders',
                'Inner borders',
                'Horizontal borders',
                'Vertical borders',
                'Left border',
                'Top border',
                'Right border',
                'Bottom border',
                'No borders',
            ]);
            expect(document.activeElement).to.equal(types[0]);

            // The arrows move in the grid of five columns, and the types are one tab stop
            key(types[0], 'ArrowRight', 39);
            expect(document.activeElement).to.equal(types[1]);
            key(types[1], 'ArrowDown', 40);
            expect(document.activeElement).to.equal(types[6]);
            key(types[6], 'ArrowDown', 40);
            expect(document.activeElement).to.equal(types[6]);
            key(types[6], 'ArrowLeft', 37);
            key(types[5], 'ArrowUp', 38);
            expect(document.activeElement).to.equal(types[0]);
            key(types[0], 'End', 35);
            expect(document.activeElement).to.equal(types[9]);
            key(types[9], 'Home', 36);
            expect(document.activeElement).to.equal(types[0]);
            key(types[0], 'ArrowLeft', 37);
            expect(document.activeElement).to.equal(types[0]);
            key(types[0], 'ArrowRight', 39);
            expect(types.map((type) => type.getAttribute('tabindex'))).to.eql(['-1', '0', ...Array(8).fill('-1')]);

            // Tab goes to the types, the border color, the thickness and the style, and stays in the dialog
            const color = dialog.querySelector<HTMLElement>('button[aria-label="Border color"]')!;
            const thickness = dialog.querySelector<HTMLElement>('[aria-label="Border thickness"]')!;
            const style = dialog.querySelector<HTMLElement>('[aria-label="Border style"]')!;
            const tabStops = Array.from(dialog.querySelectorAll<HTMLElement>('button, [tabindex]')).filter(
                (element) => element.getAttribute('tabindex') !== '-1' && !element.closest('.jcolor')
            );
            expect(tabStops).to.eql([types[1], color, thickness, style]);
            expect(thickness.getAttribute('role')).to.equal('combobox');
            expect(style.getAttribute('role')).to.equal('combobox');

            style.focus();
            expect(key(style, 'Tab', 9).defaultPrevented).to.equal(true);
            expect(document.activeElement).to.equal(types[1]);
            const shiftTab = new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true });
            types[1].dispatchEvent(shiftTab);
            expect(shiftTab.defaultPrevented).to.equal(true);
            expect(document.activeElement).to.equal(style);
            expect(key(thickness, 'Tab', 9).defaultPrevented).to.equal(false);

            // The color picker opens on the color, and Escape closes only the color picker
            color.focus();
            color.click();
            const picker = dialog.querySelector<HTMLElement>('.jcolor')!;
            expect(picker.classList.contains('jcolor-focus')).to.equal(true);
            const red = picker.querySelector<HTMLElement>('td[data-value="#f44336"]')!;
            red.focus();
            key(red, 'Enter', 13);
            expect(color.style.color).to.equal('rgb(244, 67, 54)');
            key(red, 'Escape', 27);
            expect(picker.classList.contains('jcolor-focus')).to.equal(false);
            expect(document.activeElement).to.equal(color);
            expect(dialog.hidden).to.equal(false);

            // The thickness and the style are pickers, and Escape in their lists closes only the list
            thickness.focus();
            key(thickness, 'Enter', 13);
            key(thickness, 'ArrowDown', 40);
            key(thickness, 'ArrowDown', 40);
            key(thickness, 'Enter', 13);
            style.focus();
            key(style, 'Enter', 13);
            key(style, 'ArrowDown', 40);
            key(style, 'Escape', 27);
            expect(style.getAttribute('aria-expanded')).to.equal('false');
            expect(dialog.hidden).to.equal(false);
            key(style, 'Enter', 13);
            key(style, 'ArrowDown', 40);
            key(style, 'Enter', 13);
            expect(dialog.hidden).to.equal(false);

            // A type of borders sets the borders, closes the dialog and gives the focus back to the button
            types[0].focus();
            types[0].click();
            expect(worksheet.getCell('A1').getAttribute('style')).to.contain('border-top: 3px dotted rgb(244, 67, 54)');
            expect(dialog.hidden).to.equal(true);
            expect(borders.getAttribute('aria-expanded')).to.equal('false');
            expect(document.activeElement).to.equal(borders);
            expect(borders.querySelector('i')!.textContent).to.equal('border_all');

            // The down arrow opens the dialog on the type chosen last, and Escape closes it
            expect(key(borders, 'ArrowDown', 40).defaultPrevented).to.equal(true);
            expect(document.activeElement).to.equal(types[0]);
            key(types[0], 'Escape', 27);
            expect(dialog.hidden).to.equal(true);
            expect(document.activeElement).to.equal(borders);
        });

        it('closes the borders on a press outside them, and does not open them on a worksheet that is not editable (issue #146)', async () => {
            const instance = await createWithToolbar();

            const borders = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Borders"]')!;
            borders.click();
            const dialog = document.getElementById(borders.getAttribute('aria-controls')!)!;
            expect(dialog.hidden).to.equal(false);

            // A press in the dialog or on the button keeps it open, and the button closes it
            dialog.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            borders.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(dialog.hidden).to.equal(false);
            borders.click();
            expect(dialog.hidden).to.equal(true);

            borders.click();
            document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            expect(dialog.hidden).to.equal(true);
            expect(borders.getAttribute('aria-expanded')).to.equal('false');

            instance[0].options.editable = false;
            instance[0].parent.toolbarManager.update(instance[0]);
            expect(borders.getAttribute('aria-disabled')).to.equal('true');
            borders.click();
            key(borders, 'ArrowDown', 40);
            expect(dialog.hidden).to.equal(true);
        });

        it('presses a button with Enter and Space (issue #125)', async () => {
            const instance = await createWithToolbar();
            const worksheet = instance[0];
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);

            const bold = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Bold"]')!;
            bold.focus();
            expect(key(bold, 'Enter', 13).defaultPrevented).to.equal(true);
            expect(worksheet.getCell('A1').style.fontWeight).to.equal('bold');

            key(bold, ' ', 32);
            expect(worksheet.getCell('A1').style.fontWeight).to.equal('');

            // Undo is a button too
            const undo = root.querySelector('.jss_toolbar [aria-label="Undo"]')!;
            key(undo, 'Enter', 13);
            expect(worksheet.getCell('A1').style.fontWeight).to.equal('bold');
        });

        it('names its items, and says which are disabled (issue #125)', async () => {
            const instance = await createWithToolbar();

            const names = Array.from(root.querySelectorAll('.jss_toolbar .jtoolbar-item[role="button"], .jss_toolbar .jtoolbar-item.jpicker')).map((item) =>
                item.getAttribute('aria-label')
            );
            expect(names).to.eql([
                'Undo',
                'Redo',
                'Save',
                'Font',
                'Font size',
                'Text alignment',
                'Bold',
                'Text color',
                'Background color',
                'Vertical alignment',
                'Merge the selected cells',
                'Borders',
                'Toggle Fullscreen',
            ]);

            instance[0].options.editable = false;
            instance[0].parent.toolbarManager.update(instance[0]);
            const bold = root.querySelector<HTMLElement>('.jss_toolbar [aria-label="Bold"]')!;
            expect(bold.getAttribute('aria-disabled')).to.equal('true');

            // A disabled button does nothing
            instance[0].updateSelectionFromCoords(0, 0, 0, 0);
            key(bold, 'Enter', 13);
            expect(instance[0].getCell('A1').style.fontWeight).to.equal('');
        });

        it('tells in aria-expanded that a press opened a picker', async () => {
            await createWithToolbar();

            const picker = root.querySelector<HTMLElement>('.jss_toolbar .jpicker')!;
            picker.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));

            expect(picker.getAttribute('aria-expanded')).to.equal('true');
        });
    });
});
