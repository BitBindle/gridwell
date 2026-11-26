import formula from '@jspreadsheet/formula';
import jspreadsheet from './index.js';

import './jspreadsheet.css';

window.jss = jspreadsheet;

jspreadsheet.setFormulaEngine(formula);

window.instance = jspreadsheet(document.getElementById('root'), {
    tabs: true,
    toolbar: true,
    worksheets: [
        {
            minDimensions: [6, 6],
        },
    ],
});
