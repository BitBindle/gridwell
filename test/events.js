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

const press = function (element) {
    element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
};

describe('Mouse events outside the grid', () => {
    it('keeps the selection when the press is on a child element of a jss_object control', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        const control = document.createElement('div');
        control.className = 'jss_object';
        control.innerHTML = '<span>Bold</span>';
        document.body.appendChild(control);

        press(worksheet.records[1][1].element);
        press(control.querySelector('span'));
        control.remove();

        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('takes a press on a cell of a host table with the class jss_object without an error', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 3] }] });

        const host = document.createElement('div');
        host.innerHTML = '<table><tbody><tr><td class="jss_object">Restore</td></tr></tbody></table>';
        document.body.appendChild(host);

        const errors = [];
        const listener = (event) => errors.push(String(event.error || event.message));
        window.addEventListener('error', listener);

        press(worksheet.records[1][1].element);
        press(host.querySelector('td'));

        window.removeEventListener('error', listener);
        host.remove();

        expect(errors).to.eql([]);
        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });
});
