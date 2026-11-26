import type { Library } from '../types/internal';

const lib = {
    // index.ts makes the gridwell function and puts it here
    gridwell: {} as Library,
    // The engine of gridwell.setFormulaEngine, which calculates the formulas
    formulaEngine: null as gridwell.FormulaEngine | null,
};

export default lib;
