/**
 * The clipboard of a worksheet: the text it copied last, the styles of the copied cells, the hidden textarea that puts
 * text on the system clipboard, and the dashed border around the copied cells.
 *
 * Each worksheet has one, as worksheet.clipboard. Copying and pasting read and change the cells, so the worksheet methods
 * copy and paste (src/utils/copyPaste.ts) do that, and keep what they copied here.
 */
export class Clipboard {
    /** The text of the last copy, or an empty string before the first copy. */
    data = '';

    /** The hash of data, or null before the first copy. A paste compares it with the hash of the pasted text. */
    hash: number | null = null;

    /** The styles of the copied cells, one entry for each cell, by row. */
    style: string[] = [];

    /** The hidden textarea that the copy selects, to put the text on the system clipboard. */
    textarea: HTMLTextAreaElement;

    constructor() {
        this.textarea = document.createElement('textarea');
        this.textarea.className = 'jss_textarea';
        this.textarea.id = 'jss_textarea';
        this.textarea.tabIndex = -1;
        this.textarea.ariaHidden = 'true';
    }

    /** Keep the text of a copy and the styles of its cells. */
    keep(data: string, style: string[]) {
        this.data = data;
        this.style = style;
        this.hash = hashText(data);
    }

    /** Whether text is the text of the last copy, so that a paste of it can use the data and the styles of the copy. */
    matches(text: string) {
        return hashText(text) == this.hash;
    }

    /** Put text on the system clipboard. */
    write(text: string) {
        this.textarea.value = text;
        this.textarea.select();
        document.execCommand('copy');
    }

    /** Draw the dashed border around the copied cells, from the borders of their selection. */
    showBorder(cells: { element: HTMLTableCellElement }[]) {
        for (let i = 0; i < cells.length; i++) {
            cells[i].element.classList.add('copying');
            if (cells[i].element.classList.contains('highlight-left')) {
                cells[i].element.classList.add('copying-left');
            }
            if (cells[i].element.classList.contains('highlight-right')) {
                cells[i].element.classList.add('copying-right');
            }
            if (cells[i].element.classList.contains('highlight-top')) {
                cells[i].element.classList.add('copying-top');
            }
            if (cells[i].element.classList.contains('highlight-bottom')) {
                cells[i].element.classList.add('copying-bottom');
            }
        }
    }

    /** Remove the dashed border from the cells of every worksheet in the document: there is one system clipboard. */
    clearBorder() {
        const copying = document.querySelectorAll('.jss_worksheet .copying');
        for (let i = 0; i < copying.length; i++) {
            copying[i].classList.remove('copying');
            copying[i].classList.remove('copying-left');
            copying[i].classList.remove('copying-right');
            copying[i].classList.remove('copying-top');
            copying[i].classList.remove('copying-bottom');
        }
    }
}

/** A 32-bit hash of a text. */
const hashText = function (str: string) {
    let hash = 0,
        i,
        chr;

    if (!str || str.length === 0) {
        return hash;
    } else {
        for (i = 0; i < str.length; i++) {
            chr = str.charCodeAt(i);
            hash = (hash << 5) - hash + chr;
            hash |= 0;
        }
    }
    return hash;
};
