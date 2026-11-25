const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const load = function (options, element) {
    return new Promise((resolve) => {
        const instance = jspreadsheet(element || root, {
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

describe('Keys on a jss_object control outside the grid', () => {
    const keydown = function (target, key, keyCode) {
        const event = new window.KeyboardEvent('keydown', { bubbles: true, cancelable: true, composed: true, key: key, keyCode: keyCode, which: keyCode });
        target.dispatchEvent(event);
        return event;
    };

    const paste = function (target, text) {
        const event = new window.Event('paste', { bubbles: true, cancelable: true });
        event.clipboardData = { getData: () => text };
        target.dispatchEvent(event);
        return event;
    };

    it('keeps the selection and the default action of the keys on a jss_object button', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] });

        const button = document.createElement('input');
        button.type = 'button';
        button.className = 'jss_object';
        document.body.appendChild(button);

        press(worksheet.records[1][1].element);
        press(button);
        button.focus();

        const results = [keydown(button, 'ArrowDown', 40), keydown(button, 'Enter', 13), keydown(button, 'Tab', 9)].map((event) => event.defaultPrevented);
        button.remove();

        expect(results).to.eql([false, false, false]);
        expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
    });

    it('keeps the selection with the keys on a jss_object button in a shadow root', async () => {
        const host = document.createElement('div');
        document.body.appendChild(host);
        const shadowRoot = host.attachShadow({ mode: 'open' });
        const container = document.createElement('div');
        const button = document.createElement('input');
        button.type = 'button';
        button.className = 'jss_object';
        shadowRoot.appendChild(container);
        shadowRoot.appendChild(button);

        try {
            const worksheet = await load({ root: shadowRoot, worksheets: [{ minDimensions: [3, 4] }] }, container);

            press(worksheet.records[1][1].element);
            button.focus();

            const event = keydown(button, 'ArrowDown', 40);

            expect(event.defaultPrevented).to.equal(false);
            expect(worksheet.selectedCell.map(Number)).to.eql([1, 1, 1, 1]);
        } finally {
            host.remove();
        }
    });

    it('keeps the cells with a paste on a jss_object text input', async () => {
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']] }] });

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'jss_object';
        document.body.appendChild(input);

        press(worksheet.records[0][0].element);
        press(input);
        input.focus();

        const event = paste(input, 'c');
        input.remove();

        expect(event.defaultPrevented).to.equal(false);
        expect(worksheet.getValue('A1')).to.equal('a');
    });

    it('takes the keys and a paste on a worksheet in a jss_object element', async () => {
        const wrapper = document.createElement('div');
        wrapper.className = 'jss_object';
        const container = document.createElement('div');
        wrapper.appendChild(container);
        document.body.appendChild(wrapper);

        try {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] }, container);

            press(worksheet.records[1][1].element);
            container.focus();

            const results = [keydown(container, 'ArrowDown', 40), keydown(container, 'ArrowRight', 39)].map((event) => event.defaultPrevented);
            expect(results).to.eql([true, true]);
            expect(worksheet.selectedCell.map(Number)).to.eql([2, 2, 2, 2]);

            paste(container, 'c');
            expect(worksheet.getValue('C3')).to.equal('c');
        } finally {
            wrapper.remove();
        }
    });

    it('takes the keys and a paste on an element without the class jss_object', async () => {
        const worksheet = await load({ worksheets: [{ minDimensions: [3, 4] }] });

        press(worksheet.records[1][1].element);

        const event = keydown(document.body, 'ArrowDown', 40);
        expect(event.defaultPrevented).to.equal(true);
        expect(worksheet.selectedCell.map(Number)).to.eql([1, 2, 1, 2]);

        paste(document.body, 'c');
        expect(worksheet.getValue('B3')).to.equal('c');
    });

    it('closes a custom editor with a jss_object control with Escape', async () => {
        // A custom editor with a jss_object control outside the worksheet, for example a floating panel
        const control = document.createElement('input');
        control.className = 'jss_object';
        const editorType = {
            createCell: function (cell, value) {
                cell.textContent = value;
                return cell;
            },
            updateCell: function (cell, value) {
                cell.textContent = value;
                return value;
            },
            openEditor: function (cell) {
                cell.classList.add('editor');
                document.body.appendChild(control);
                control.focus();
            },
            closeEditor: function () {
                control.remove();
                return control.value;
            },
        };
        const worksheet = await load({ worksheets: [{ data: [['a', 'b']], columns: [{}, { type: editorType }] }] });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        worksheet.openEditor(worksheet.records[0][1].element);
        control.value = 'c';

        keydown(control, 'Escape', 27);
        control.remove();

        expect(worksheet.edition).to.equal(null);
        expect(worksheet.getValue('B1')).to.equal('b');
    });
});

describe('Dispatch events', () => {
    it('gives onevent and plugins the event name and the arguments, and the event handler the arguments', () => {
        const calls = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a', 'b']] }],
            onevent: function (event, instance, cell, x, y, value) {
                if (event === 'onchange') {
                    calls.push(['onevent', instance === worksheet, String(x), String(y), value]);
                }
            },
            onchange: function (instance, cell, x, y, value) {
                calls.push(['onchange', instance === worksheet, String(x), String(y), value]);
            },
            plugins: {
                test: function () {
                    return {
                        onevent: function (event, instance, cell, x, y, value) {
                            if (event === 'onchange') {
                                calls.push(['plugin', instance === worksheet, String(x), String(y), value]);
                            }
                        },
                    };
                },
            },
        })[0];

        worksheet.setValueFromCoords(1, 0, 'c');

        expect(calls).to.eql([
            ['onevent', true, '1', '0', 'c'],
            ['onchange', true, '1', '0', 'c'],
            ['plugin', true, '1', '0', 'c'],
        ]);
    });

    it('returns the result of the handler, so that onbeforeinsertrow can cancel the insertion', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            onbeforeinsertrow: function () {
                return false;
            },
        })[0];

        worksheet.insertRow();

        expect(worksheet.getData()).to.eql([['a']]);
    });
});
