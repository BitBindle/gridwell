import type { Library } from '../types/internal';

const lib = {
    // index.ts makes the jspreadsheet function and puts it here
    jspreadsheet: {} as Library,
    // The engine of jspreadsheet.setFormulaEngine, which calculates the formulas
    formulaEngine: null as jspreadsheet.FormulaEngine | null,
};

export default lib;
