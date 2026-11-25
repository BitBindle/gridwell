const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const load = function (options) {
    return new Promise((resolve) => {
        const instance = jspreadsheet(root, {
            ...options,
            onload: () => resolve(instance[0]),
        });
    });
};

// A short touch: touchend stops the timer of the long touch
const tap = function (element) {
    element.dispatchEvent(new window.Event('touchstart', { bubbles: true }));
    element.dispatchEvent(new window.Event('touchend', { bubbles: true }));
};

describe('Touch controls', () => {
    let control;

    afterEach(() => {
        if (control) {
            control.remove();
            control = null;
        }
    });

    it('keeps the selection when a tap is on a jss_object control', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });
        control = document.body.appendChild(document.createElement('button'));
        control.className = 'jss_object';

        tap(worksheet.records[1][1].element);
        tap(control);

        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('keeps the selection when a tap is on a child element of a jss_object control', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });
        control = document.body.appendChild(document.createElement('div'));
        control.className = 'jss_object';
        control.innerHTML = '<button><span>Bold</span></button>';

        tap(worksheet.records[1][1].element);
        tap(control.querySelector('span'));

        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('selects no cell of the worksheet when a tap is on a cell of a jss_object table', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });
        control = document.body.appendChild(document.createElement('table'));
        control.className = 'jss_object';
        control.innerHTML = '<tbody><tr><td data-x="2" data-y="2">x</td></tr></tbody>';

        tap(worksheet.records[1][1].element);
        tap(control.querySelector('td'));

        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('resets the selection when a tap is outside the worksheet and outside a jss_object control', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });
        control = document.body.appendChild(document.createElement('button'));

        tap(worksheet.records[1][1].element);
        tap(control);

        expect(worksheet.selectedCell).to.equal(null);
    });
});
