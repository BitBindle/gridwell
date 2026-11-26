import type { TargetEvent, Worksheet } from '../types/internal';

/**
 * The index in the data-x or data-y attribute of an element, or null when the element has none
 */
export function getIndex(element: Element, name: 'data-x' | 'data-y') {
    const text = element.getAttribute(name);
    return text ? parseInt(text) : null;
}

// The element of the worksheet that contains the element, which keeps the worksheet, and the section: 1 for the header, 2 for the body
export function getElement(element: HTMLElement): [Worksheet['element'] | 0 | null, number] {
    let jssSection = 0;
    let jssElement: Worksheet['element'] | 0 | null = 0;

    // The path goes up to the document, which has no className
    function path(element: HTMLElement) {
        if (element.className) {
            if (element.classList.contains('jss_container')) {
                jssElement = element as Worksheet['element'];
            }

            if (element.classList.contains('jss_spreadsheet')) {
                jssElement = element.querySelector(':scope > .jtabs-content > .jtabs-selected') as Worksheet['element'] | null;
            }
        }

        if (element.tagName == 'THEAD') {
            jssSection = 1;
        } else if (element.tagName == 'TBODY') {
            jssSection = 2;
        }

        if (element.parentNode) {
            if (!jssElement) {
                path(element.parentNode as HTMLElement);
            }
        }
    }

    path(element);

    // A section of a table outside a worksheet is not a section of the grid
    return [jssElement, jssElement ? jssSection : 0];
}

export function getRole(element: HTMLElement) {
    if (element.classList.contains('jss_selectall')) {
        return 'select-all';
    }

    if (element.classList.contains('jss_corner')) {
        return 'fill-handle';
    }

    let tempElement = element;

    while (!tempElement.classList.contains('jss_spreadsheet')) {
        if (tempElement.classList.contains('jss_row')) {
            return 'row';
        }

        if (tempElement.classList.contains('jss_nested')) {
            return 'nested';
        }

        if (tempElement.classList.contains('jtabs-headers')) {
            return 'tabs';
        }

        if (tempElement.classList.contains('jtoolbar')) {
            return 'toolbar';
        }

        if (tempElement.classList.contains('jss_pagination')) {
            return 'pagination';
        }

        if (tempElement.tagName === 'TBODY') {
            return 'cell';
        }

        if (tempElement.tagName === 'TFOOT') {
            return getElementIndex(element) === 0 ? 'grid' : 'footer';
        }

        if (tempElement.tagName === 'THEAD') {
            return 'header';
        }

        // The element is in a worksheet, so the loop ends at the element of the spreadsheet
        tempElement = tempElement.parentElement!;
    }

    return 'applications';
}

export function getElementIndex(element: HTMLElement) {
    const parentChildren = element.parentElement!.children;

    for (let i = 0; i < parentChildren.length; i++) {
        const currentElement = parentChildren[i];

        if (element === currentElement) {
            return i;
        }
    }

    return -1;
}

function isMac() {
    return navigator.platform.toUpperCase().includes('MAC');
}

export function isCtrl(e: TargetEvent<KeyboardEvent | MouseEvent>) {
    if (isMac()) {
        return e.metaKey;
    } else {
        return e.ctrlKey;
    }
}
