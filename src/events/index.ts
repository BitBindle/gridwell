import { pasteControls } from './clipboard.js';
import { contextMenuControls } from './contextMenu.js';
import { focusInControls } from './focus.js';
import { keyDownControls } from './keyboard.js';
import { doubleClickControls, mouseDownControls, mouseMoveControls, mouseOverControls, mouseUpControls } from './mouse.js';
import { scrollControls } from './scroll.js';
import { touchEndControls, touchStartControls } from './touch.js';

export function setEvents(root: Document | HTMLElement | ShadowRoot) {
    destroyEvents(root);
    // The handlers read the target as an element, which the type of a listener does not promise
    root.addEventListener('mouseup', mouseUpControls as EventListener);
    root.addEventListener('mousedown', mouseDownControls as EventListener);
    root.addEventListener('mousemove', mouseMoveControls as EventListener);
    root.addEventListener('mouseover', mouseOverControls as EventListener);
    root.addEventListener('dblclick', doubleClickControls as EventListener);
    root.addEventListener('paste', pasteControls as EventListener);
    root.addEventListener('contextmenu', contextMenuControls as EventListener);
    root.addEventListener('touchstart', touchStartControls as EventListener);
    root.addEventListener('touchend', touchEndControls as EventListener);
    root.addEventListener('touchcancel', touchEndControls as EventListener);
    root.addEventListener('touchmove', touchEndControls as EventListener);
    root.addEventListener('focusin', focusInControls as EventListener);
    document.addEventListener('keydown', keyDownControls as EventListener);
    // A scroll does not bubble, so the document sees the scroll of the page and of any element in the capture phase
    document.addEventListener('scroll', scrollControls, true);
    // A scroll in a shadow root does not leave it, so each shadow root around the spreadsheet listens too
    for (const shadow of shadowsOf(root)) {
        shadow.addEventListener('scroll', scrollControls, true);
    }
}

/**
 * The shadow roots that hold the root, from the nearest one to the one in the document. A root that is a shadow root
 * comes first.
 */
function shadowsOf(root: Document | HTMLElement | ShadowRoot) {
    const shadows: ShadowRoot[] = [];
    let node = root.getRootNode() as Node;
    while (node !== document && (node as ShadowRoot).host) {
        shadows.push(node as ShadowRoot);
        node = (node as ShadowRoot).host.getRootNode();
    }
    return shadows;
}

export function destroyEvents(root: Document | HTMLElement | ShadowRoot) {
    // The handlers read the target as an element, which the type of a listener does not promise
    root.removeEventListener('mouseup', mouseUpControls as EventListener);
    root.removeEventListener('mousedown', mouseDownControls as EventListener);
    root.removeEventListener('mousemove', mouseMoveControls as EventListener);
    root.removeEventListener('mouseover', mouseOverControls as EventListener);
    root.removeEventListener('dblclick', doubleClickControls as EventListener);
    root.removeEventListener('paste', pasteControls as EventListener);
    root.removeEventListener('contextmenu', contextMenuControls as EventListener);
    root.removeEventListener('touchstart', touchStartControls as EventListener);
    root.removeEventListener('touchend', touchEndControls as EventListener);
    root.removeEventListener('touchcancel', touchEndControls as EventListener);
    root.removeEventListener('touchmove', touchEndControls as EventListener);
    root.removeEventListener('focusin', focusInControls as EventListener);
    document.removeEventListener('keydown', keyDownControls as EventListener);
    document.removeEventListener('scroll', scrollControls, true);
    for (const shadow of shadowsOf(root)) {
        shadow.removeEventListener('scroll', scrollControls, true);
    }
}
