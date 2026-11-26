/* eslint-disable no-var -- a property of globalThis is declared with var */

// The globals of the tests. They run in Node, but declare only these instead of the Node types, whose setTimeout returns
// an object instead of the number that the code of the library keeps.

// The element that mocha.config.js adds to the document for each test. It is emptied after each test.
declare var root: HTMLDivElement;

// The global object of Node, where jsdom-global puts window, document and the other browser globals
declare var global: typeof globalThis;

declare function require(id: string): unknown;
