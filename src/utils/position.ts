/** The side of an element where its popup opens, and the room on that side. */
export interface PopupSide {
    /** The popup opens over the element. */
    above: boolean;
    /** The height in pixels that the popup has on that side. */
    room: number;
}

/**
 * Tell on which side of an element its popup opens: under it, or over it when the popup does not fit under it and there
 * is more room over it. The room ends at the edges of the window and, for a popup without a fixed position, at the edges
 * of the ancestors that clip it, such as a container that scrolls.
 * @param anchor - The element that the popup opens from.
 * @param popup - The popup, an element with an absolute or a fixed position.
 * @param fixed - The popup has a fixed position, so that no ancestor clips it.
 */
export function getPopupSide(anchor: HTMLElement, popup: HTMLElement, fixed: boolean): PopupSide {
    const rect = anchor.getBoundingClientRect();
    const height = popup.getBoundingClientRect().height;
    let top = 0;
    let bottom = window.innerHeight;

    if (!fixed) {
        // An ancestor clips the popup when it holds the containing block of the popup: the first ancestor with a position
        let holdsContainingBlock = false;
        for (let element = popup.parentElement; element && element !== document.body; element = element.parentElement) {
            const style = getComputedStyle(element);
            if (!holdsContainingBlock && style.position !== 'static') {
                holdsContainingBlock = true;
            }
            if (holdsContainingBlock && style.overflowY !== 'visible') {
                const bounds = element.getBoundingClientRect();
                top = Math.max(top, bounds.top);
                bottom = Math.min(bottom, bounds.bottom);
            }
        }
    }

    const below = bottom - rect.bottom;
    const over = rect.top - top;
    const above = height > below && over > below;
    return { above, room: Math.max(0, above ? over : below) };
}
