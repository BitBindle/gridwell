/**
 * Global declarations for the type check of the JavaScript source. They are not published.
 */

/**
 * The code calls parseInt with numbers in many places, for example with coordinates and widths. JavaScript converts the
 * number to a string first. This overload allows it until those calls are changed, so the type check can report other
 * problems first.
 */
declare function parseInt(value: string | number, radix?: number): number;

interface Navigator {
    /** Internet Explorer and the old Edge only. download() uses it when it is there. */
    msSaveOrOpenBlob?: (blob: Blob, defaultName?: string) => boolean;
}

interface Element {
    /**
     * The code sets attributes to numbers in many places, for example colspan and data-x. The DOM converts the number
     * to a string.
     */
    setAttribute(qualifiedName: string, value: string | number): void;
}
