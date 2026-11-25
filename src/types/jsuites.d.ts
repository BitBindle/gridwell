/**
 * jsuites has these methods, but its typings do not declare them.
 */

// This file must be a module, so that the declaration below adds to the JSuites interface instead of replacing the module
export {};

declare module 'jsuites/dist/types' {
    interface JSuites {
        /** Translate a text with the dictionary from setDictionary. */
        translate(text: string): string;

        /** Make a random unique id. */
        guid(): string;

        /** Make an image upload control in the element. */
        image(element: HTMLElement, options?: object): unknown;

        /** Make one request, or several requests and call complete when all have finished. */
        ajax(options: object | object[], complete?: () => void): unknown;
    }
}
