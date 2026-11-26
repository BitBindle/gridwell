/** The tags that the rich text keeps. Another tag is replaced by its text, and a script or a style is removed. */
const allowedTags = new Set([
    'a',
    'b',
    'blockquote',
    'br',
    'code',
    'del',
    'div',
    'em',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'hr',
    'i',
    'ins',
    'li',
    'mark',
    'ol',
    'p',
    'pre',
    's',
    'small',
    'span',
    'strong',
    'sub',
    'sup',
    'u',
    'ul',
]);

/** The tags whose content is code, not text. */
const removedTags = new Set(['script', 'style', 'template', 'iframe', 'object', 'embed', 'noscript']);

/** The properties of the style attribute that the rich text keeps. */
const allowedStyles = new Set(['color', 'background-color', 'font-weight', 'font-style', 'font-size', 'text-decoration', 'text-align']);

/** The protocols of a link that the rich text keeps. */
const allowedProtocols = /^(https?:|mailto:|tel:|#|\/)/i;

/**
 * Keep the tags of formatting, the address of a link and some properties of the style, and remove everything else:
 * scripts, event handlers, forms, images and frames. It is the HTML that a rich text editor saves and pastes.
 * @param html - The HTML.
 */
export function sanitizeHtml(html: string) {
    // A template parses the HTML without loading its images or running its scripts
    const template = document.createElement('template');
    template.innerHTML = html;
    clean(template.content);
    return template.innerHTML;
}

/**
 * Clean the children of a node.
 * @param node - The node.
 */
function clean(node: Node) {
    for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === Node.ELEMENT_NODE) {
            const element = child as HTMLElement;
            const tag = element.tagName.toLowerCase();
            if (removedTags.has(tag)) {
                element.remove();
            } else if (!allowedTags.has(tag)) {
                // The text of the element stays, without its tags
                clean(element);
                element.replaceWith(...Array.from(element.childNodes));
            } else {
                cleanAttributes(element, tag);
                clean(element);
            }
        } else if (child.nodeType !== Node.TEXT_NODE) {
            // A comment, or a processing instruction
            child.parentNode!.removeChild(child);
        }
    }
}

/**
 * Remove the attributes of an element, except the address of a link and the properties of the style that the rich text
 * keeps.
 * @param element - The element.
 * @param tag - Its tag, in lower case.
 */
function cleanAttributes(element: HTMLElement, tag: string) {
    const href = element.getAttribute('href');
    const styles: string[] = [];
    for (const declaration of (element.getAttribute('style') || '').split(';')) {
        const index = declaration.indexOf(':');
        const property = declaration.slice(0, index).trim().toLowerCase();
        const value = declaration.slice(index + 1).trim();
        if (index > 0 && allowedStyles.has(property) && !/[()\\<>]/.test(value)) {
            styles.push(`${property}: ${value}`);
        }
    }

    for (const name of element.getAttributeNames()) {
        element.removeAttribute(name);
    }

    if (tag === 'a' && href && allowedProtocols.test(href.trim())) {
        element.setAttribute('href', href.trim());
    }
    if (styles.length) {
        element.setAttribute('style', styles.join('; '));
    }
}

/**
 * Make an element a rich text editor: a text box that takes formatted text, such as bold with Ctrl+B, and that pastes
 * the HTML of the clipboard without its scripts and styles.
 * @param element - The element.
 * @param value - The HTML that the editor starts with.
 * @param label - The accessible name of the editor.
 */
export function createRichText(element: HTMLElement, value: string, label: string) {
    element.setAttribute('contenteditable', 'true');
    element.setAttribute('role', 'textbox');
    element.setAttribute('aria-multiline', 'true');
    element.setAttribute('aria-label', label);
    element.innerHTML = sanitizeHtml(value);

    element.addEventListener('paste', (e) => {
        const html = e.clipboardData?.getData('text/html');
        const text = e.clipboardData?.getData('text/plain');
        if (!html && !text) {
            return;
        }
        e.preventDefault();

        const fragment = document.createElement('template');
        if (html) {
            fragment.innerHTML = sanitizeHtml(html);
        } else {
            // Each line of the text is a line of the editor
            text!.split(/\r?\n/).forEach((line, index) => {
                if (index > 0) {
                    fragment.content.appendChild(document.createElement('br'));
                }
                fragment.content.appendChild(document.createTextNode(line));
            });
        }
        insertAtCaret(element, fragment.content);
    });

    // The caret goes after the text
    element.focus();
    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
}

/**
 * Put nodes where the caret of an editor is, in place of the selected text, and move the caret after them.
 * @param element - The editor.
 * @param content - The nodes.
 */
function insertAtCaret(element: HTMLElement, content: DocumentFragment) {
    const selection = window.getSelection();
    let range = selection?.rangeCount ? selection.getRangeAt(0) : null;
    if (!range || !element.contains(range.commonAncestorContainer)) {
        range = document.createRange();
        range.selectNodeContents(element);
        range.collapse(false);
    }

    const last = content.lastChild;
    range.deleteContents();
    range.insertNode(content);
    if (last) {
        range.setStartAfter(last);
        range.collapse(true);
        selection?.removeAllRanges();
        selection?.addRange(range);
    }
}

/**
 * The HTML of a rich text editor, without what sanitizeHtml removes.
 * @param element - The editor.
 */
export function getRichText(element: HTMLElement) {
    return sanitizeHtml(element.innerHTML);
}
