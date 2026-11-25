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
