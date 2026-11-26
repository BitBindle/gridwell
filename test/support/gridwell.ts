/**
 * The library for the tests: the built bundle, which the tests run, with the public typings. A worksheet that it
 * returns has the type of the class, so that a test can read the state that only the library uses, such as
 * worksheet.selection.range.
 */
import formula from '@jspreadsheet/formula';

import type Gridwell from '../../src/gridwell';
import type { Worksheet } from '../../src/types/internal';

type Library = {
    (element: HTMLTableElement, options?: Gridwell.TableSpreadsheetOptions): Worksheet[];
    (element: HTMLElement, options: Gridwell.SpreadsheetOptions): Worksheet[];
} & Gridwell.Gridwell;

// The bundle is a UMD module, so it is required. The tests import the bundle and not src/, which the build makes.
const gridwell = require('../../dist/index.js') as Library;

// The tests calculate the formulas with the engine that applications register
gridwell.setFormulaEngine(formula);

export default gridwell;

export type { Gridwell, Worksheet };
