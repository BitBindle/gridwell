/**
 * The jsuites module for the type check, with the typings of jsuites corrected where they do not match what jsuites does.
 *
 * tsconfig.json maps the module 'jsuites' to this file. The build still bundles nothing from jsuites: it gets jsuites from
 * a global or require (see build.cjs).
 */
import type { JSuites as Base } from 'jsuites/dist/types';

/** The options of a jsuites component, the second parameter of the function that creates it */
type ComponentOptions<F> = F extends (el: HTMLElement, options: infer O) => unknown ? O : never;

/**
 * The options of jSuites.calendar. jsuites has defaults for each of them, takes a truthy position for a fixed position,
 * and gives onclose the input.
 */
export type CalendarOptions = Partial<Omit<ComponentOptions<Base['calendar']>, 'onclose' | 'position'>> & {
    position?: boolean;
    onclose?: (el: HTMLElement) => void;
};

export interface JSuites extends Omit<Base, 'calendar' | 'dropdown'> {
    calendar(el: HTMLElement, options: CalendarOptions): ReturnType<Base['calendar']>;

    /** Translate a text with the dictionary from setDictionary. */
    translate(text: string): string;

    /** Make a random unique id. */
    guid(): string;

    /** Make an image upload control in the element. */
    image(element: HTMLElement, options?: object): unknown;

    /** Make one request, or several requests and call complete when all have finished. */
    ajax(options: object | object[], complete?: () => void): unknown;
}

declare const jSuites: JSuites;

export default jSuites;
