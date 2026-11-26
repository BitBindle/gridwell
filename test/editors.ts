import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

const codes: Record<string, number> = { Tab: 9, Enter: 13, Escape: 27, ' ': 32 };

const key = (element: Element, keyName: string, options: KeyboardEventInit = {}) => {
    const keyCode = codes[keyName];
    const event = new KeyboardEvent('keydown', { key: keyName, keyCode, which: keyCode, bubbles: true, cancelable: true, ...options });
    element.dispatchEvent(event);
    return event;
};

const waitForListeners = () => new Promise((resolve) => setTimeout(resolve));

/** A paste event with the texts of a clipboard, which jsdom cannot make. */
const paste = (element: Element, data: Record<string, string>) => {
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: { getData: (type: string) => data[type] ?? '' } });
    element.dispatchEvent(event);
    return event;
};

describe('The rich text editor of an html column', () => {
    const option = global.Option;

    beforeEach(() => {
        // The html cells use the Option constructor of the window
        global.Option = window.Option;
    });

    afterEach(() => {
        global.Option = option;
    });

    const open = async (value: string) => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [[value, 'x']], columns: [{ type: 'html' }, {}] }] })[0];
        await waitForListeners();
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        const cell = worksheet.getCell('A1');
        worksheet.openEditor(cell);
        const editor = cell.querySelector<HTMLElement>('.jss_richtext')!;
        return { worksheet, cell, editor };
    };

    it('starts with the HTML of the cell, without its scripts, and has the focus', async () => {
        const { editor } = await open('<b>bold</b><script>alert(1)</script><img src="x" onerror="alert(1)">');

        expect(editor.getAttribute('contenteditable')).to.equal('true');
        expect(editor.innerHTML).to.equal('<b>bold</b>');
        expect(document.activeElement).to.equal(editor);
    });

    it('saves the HTML with Tab, and the selection moves right', async () => {
        const { worksheet, editor } = await open('a');

        editor.innerHTML = '<i>new</i>';
        key(editor, 'Tab');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('<i>new</i>');
        expect(worksheet.getCell('A1').innerHTML).to.equal('<i>new</i>');
        expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);
    });

    it('keeps the editor open with Enter, which adds a line', async () => {
        const { worksheet, editor } = await open('a');

        key(editor, 'Enter');

        expect(worksheet.editor.isOpen()).to.equal(true);
    });

    it('closes without saving with Escape', async () => {
        const { worksheet, editor } = await open('a');

        editor.innerHTML = 'b';
        key(editor, 'Escape');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('a');
    });

    it('saves only the tags of formatting, the links to web addresses and some properties of the style', async () => {
        const { worksheet, editor } = await open('a');

        editor.innerHTML = [
            '<p onclick="alert(1)" style="color: red; position: fixed; background-color: url(x)">a</p>',
            '<a href="javascript:alert(1)">b</a>',
            '<a href=" https://example.com " target="_blank">c</a>',
            '<font face="Arial">d</font>',
            '<!-- e -->',
            '<style>p { color: red }</style>',
        ].join('');
        worksheet.closeEditor(worksheet.getCell('A1'), true);

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('<p style="color: red">a</p><a>b</a><a href="https://example.com">c</a>d');
    });

    it('pastes the HTML of the clipboard without its scripts and styles', async () => {
        const { worksheet, editor } = await open('');

        const event = paste(editor, { 'text/html': '<b style="font-family: Arial">b</b><script>alert(1)</script>', 'text/plain': 'b' });

        expect(event.defaultPrevented).to.equal(true);
        expect(editor.innerHTML).to.equal('<b>b</b>');
        // The paste does not go to the cells
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('x');
    });

    it('pastes a text of the clipboard as lines, where the caret is', async () => {
        const { editor } = await open('<b>a</b>');

        paste(editor, { 'text/plain': 'b\nc' });
        paste(editor, { 'text/plain': 'd' });

        expect(editor.innerHTML).to.equal('<b>a</b>b<br>cd');
    });

    it('leaves a paste without text to the browser', async () => {
        const { editor } = await open('a');

        const event = paste(editor, {});

        expect(event.defaultPrevented).to.equal(false);
        expect(editor.innerHTML).to.equal('a');
    });
});

describe('The image editor of an image column', () => {
    const image = 'data:image/png;base64,iVBORw0KGgo=';

    const open = async (value: string) => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [[value, 'x']], columns: [{ type: 'image' }, {}] }] })[0];
        await waitForListeners();
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        const cell = worksheet.getCell('A1');
        worksheet.openEditor(cell);
        const dialog = cell.querySelector<HTMLElement>('.jupload')!;
        const buttons = Array.from(dialog.querySelectorAll('button'));
        const input = dialog.querySelector<HTMLInputElement>('input[type="file"]')!;
        return { worksheet, cell, dialog, buttons, input };
    };

    /** Give the file input a file, which jsdom does not let a test choose. */
    const choose = (input: HTMLInputElement, file: File) => {
        Object.defineProperty(input, 'files', { value: [file], configurable: true });
        input.dispatchEvent(new Event('change'));
    };

    /** Wait until the dialog shows an image, which the FileReader reads after a moment, or for a second at most. */
    const loaded = (dialog: HTMLElement) =>
        new Promise<void>((resolve, reject) => {
            let tries = 0;
            const check = () => {
                if (dialog.querySelector('img')) {
                    resolve();
                } else if (++tries > 200) {
                    reject(new Error('The image was not shown'));
                } else {
                    setTimeout(check, 5);
                }
            };
            check();
        });

    it('is a dialog with the image of the cell and its buttons, and the first button has the focus', async () => {
        const { dialog, buttons } = await open(image);

        expect(dialog.getAttribute('role')).to.equal('dialog');
        expect(dialog.getAttribute('aria-label')).to.equal('Edit cell A1');
        expect(dialog.querySelector('img')!.getAttribute('src')).to.equal(image);
        expect(buttons.map((button) => button.textContent)).to.eql(['Choose an image', 'Remove', 'Done']);
        expect(document.activeElement).to.equal(buttons[0]);
    });

    it('removes the image, and saves with Done', async () => {
        const { worksheet, cell, dialog, buttons } = await open(image);

        buttons[1].click();
        expect(dialog.querySelector('img')).to.equal(null);
        expect(buttons[1].disabled).to.equal(true);
        buttons[2].click();

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('');
        expect(cell.querySelector('img')).to.equal(null);
    });

    it('keeps the focus in the dialog with Tab, past the disabled Remove', async () => {
        const { worksheet, buttons } = await open('');

        key(buttons[0], 'Tab');
        expect(document.activeElement).to.equal(buttons[2]);
        key(buttons[2], 'Tab');
        expect(document.activeElement).to.equal(buttons[0]);
        key(buttons[0], 'Tab', { shiftKey: true });
        expect(document.activeElement).to.equal(buttons[2]);
        expect(worksheet.editor.isOpen()).to.equal(true);
    });

    it('presses the button with the focus with Enter or Space, and Done keeps the cell selected', async () => {
        const { worksheet, buttons, input } = await open(image);

        let opened = 0;
        input.click = () => {
            opened++;
        };
        key(buttons[0], ' ');
        expect(opened).to.equal(1);

        key(buttons[1], 'Enter');
        key(buttons[2], 'Enter');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('');
        expect(worksheet.selection.range).to.eql([0, 0, 0, 0]);
    });

    it('closes without saving with Escape', async () => {
        const { worksheet, buttons } = await open(image);

        buttons[1].click();
        key(buttons[0], 'Escape');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal(image);
    });

    it('shows and saves the image file that is chosen', async () => {
        const { worksheet, cell, dialog, buttons, input } = await open('');

        choose(input, new File(['x'], 'a.png', { type: 'image/png' }));
        await loaded(dialog);
        buttons[2].click();

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('data:image/png;base64,eA==');
        expect(cell.querySelector('img')!.getAttribute('src')).to.equal('data:image/png;base64,eA==');
    });

    it('shows the image file that is dropped on it, and ignores a file that is not an image', async () => {
        const { dialog } = await open('');

        const drop = (file: File) => {
            const event = new Event('drop', { bubbles: true, cancelable: true });
            Object.defineProperty(event, 'dataTransfer', { value: { files: [file] } });
            dialog.dispatchEvent(event);
            return event;
        };

        const dragover = new Event('dragover', { bubbles: true, cancelable: true });
        dialog.dispatchEvent(dragover);
        expect(dragover.defaultPrevented).to.equal(true);

        expect(drop(new File(['x'], 'a.txt', { type: 'text/plain' })).defaultPrevented).to.equal(true);
        await new Promise((resolve) => setTimeout(resolve, 20));
        expect(dialog.querySelector('img')).to.equal(null);

        drop(new File(['x'], 'a.png', { type: 'image/png' }));
        await loaded(dialog);
        expect(dialog.querySelector('img')!.getAttribute('src')).to.equal('data:image/png;base64,eA==');
    });
});
