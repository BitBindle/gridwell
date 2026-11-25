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
});
