const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Toolbar', () => {
    let getContext;

    before(() => {
        // The color picker of the default toolbar draws on a canvas, which jsdom does not implement
        getContext = HTMLCanvasElement.prototype.getContext;
        HTMLCanvasElement.prototype.getContext = function () {
            return new Proxy(
                {},
                {
                    get: (target, name) => (name in target ? target[name] : () => ({ data: [0, 0, 0, 0], addColorStop() {} })),
                    set: (target, name, value) => {
                        target[name] = value;
                        return true;
                    },
                }
            );
        };
    });

    after(() => {
        HTMLCanvasElement.prototype.getContext = getContext;
    });

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

            const fonts = root.querySelector('.jss_toolbar .jpicker').querySelectorAll('.jpicker-item');
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
        const getIconItem = (icon) =>
            Array.from(root.querySelectorAll('.jss_toolbar .jtoolbar-item')).find((item) => item.children[0] && item.children[0].textContent === icon);

        // Choose the option at index in the picker at position among the select items
        const choose = (position, index) => {
            const picker = root.querySelectorAll('.jss_toolbar .jpicker')[position];
            picker.querySelectorAll('.jpicker-item')[index].click();
        };

        const getStyle = (worksheet, cellName) => worksheet.getCell(cellName).getAttribute('style') || '';

        it('has the items of the default toolbar, in order', async () => {
            let items = null;
            jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3] }],
                toolbar: function (toolbar) {
                    items = toolbar.items.map((item) => item.content || item.type);
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
                'select',
                'divisor',
                'fullscreen',
            ]);
        });

        it('the size, the alignment, bold and the vertical alignment style the selected cells', async () => {
            const changes = [];
            jspreadsheet(root, {
                toolbar: true,
                worksheets: [{ minDimensions: [3, 3] }],
                onchangestyle: (worksheet, newValue) => {
                    changes.push(newValue);
                },
            });
            await new Promise((resolve) => setTimeout(resolve, 0));
            const worksheet = jspreadsheet.current;
            worksheet.updateSelectionFromCoords(0, 0, 1, 1);

            choose(1, 3);
            choose(2, 1);
            getIconItem('format_bold').click();
            choose(3, 0);

            // The toolbar gives the values with a space after the colon. A browser trims it, but jsdom drops a font size
            // that starts with a space and keeps the space of the others, so the test reads the styles from the events.
            expect(changes.map((change) => Object.keys(change).sort())).to.eql([
                ['A1', 'A2', 'B1', 'B2'],
                ['A1', 'A2', 'B1', 'B2'],
                ['A1', 'A2', 'B1', 'B2'],
                ['A1', 'A2', 'B1', 'B2'],
            ]);
            ['A1', 'B1', 'A2', 'B2'].forEach((cellName) => {
                expect(changes[0][cellName], cellName).to.equal('font-size: large');
                expect(changes[1][cellName], cellName).to.equal('text-align: center');
                expect(changes[2][cellName], cellName).to.equal('font-weight:bold');
                expect(changes[3][cellName], cellName).to.equal('vertical-align: top');

                const style = getStyle(worksheet, cellName);
                expect(style, cellName).to.match(/text-align: +center/);
                expect(style, cellName).to.contain('font-weight: bold');
                expect(style, cellName).to.match(/vertical-align: +top/);
            });
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

            // border_outer
            choose(4, 1);
            expect(getStyle(worksheet, 'A1')).to.contain('border-top: 1px solid black').and.to.contain('border-left: 1px solid black');
            expect(getStyle(worksheet, 'A1')).not.to.contain('border-right: 1px').and.not.to.contain('border-bottom: 1px');
            expect(getStyle(worksheet, 'B2')).to.contain('border-right: 1px solid black').and.to.contain('border-bottom: 1px solid black');
            expect(getStyle(worksheet, 'B2')).not.to.contain('border-top: 1px').and.not.to.contain('border-left: 1px');

            // border_inner
            choose(4, 2);
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
            const calls = [];
            jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3] }],
                toolbar: function (toolbar) {
                    calls.push(['option', Array.isArray(toolbar.items)]);
                    toolbar.items = toolbar.items.filter((item) => item.content === 'undo' || item.content === 'redo');
                    return toolbar;
                },
                plugins: {
                    test: function () {
                        return {
                            toolbar: function (toolbar) {
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
});
