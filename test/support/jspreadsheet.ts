/**
 * The library for the tests: the built bundle, which the tests run, with the public typings. A worksheet that it
 * returns has the type of the class, so that a test can read the state that only the library uses, such as
 * worksheet.selection.range.
 */
import type Jspreadsheet from '../../src/jspreadsheet';
import type { Worksheet } from '../../src/types/internal';

type Library = ((element: HTMLElement, options: Jspreadsheet.SpreadsheetOptions) => Worksheet[]) & Jspreadsheet.JSpreadsheet;

// The bundle is a UMD module, so it is required. The tests import the bundle and not src/, which the build makes.
const jspreadsheet = require('../../dist/index.js') as Library;

export default jspreadsheet;

export type { Jspreadsheet, Worksheet };
