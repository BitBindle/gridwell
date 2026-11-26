/**
 * The jsuites module for the type check, with the typings of jsuites corrected where they do not match what jsuites does.
 *
 * tsconfig.json maps the module 'jsuites' to this file. The build still bundles nothing from jsuites: it gets jsuites from
 * a global or require (see build.cjs).
 */
import type { JSuites as Base } from 'jsuites/dist/types';

export interface JSuites extends Omit<Base, 'calendar' | 'dropdown' | 'editor'> {
    /** Translate a text with the dictionary from setDictionary. */
    translate(text: string): string;

    /** Make a random unique id. */
    guid(): string;

    /** Make one request, or several requests and call complete when all have finished. */
    ajax(options: object | object[], complete?: () => void): unknown;
}

declare const jSuites: JSuites;

export default jSuites;
