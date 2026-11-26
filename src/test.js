import formula from '@jspreadsheet/formula';
import gridwell from './index.js';

import './gridwell.css';

window.jss = gridwell;

gridwell.setFormulaEngine(formula);

window.instance = gridwell(document.getElementById('root'), {
    tabs: true,
    toolbar: true,
    worksheets: [
        {
            minDimensions: [6, 6],
        },
    ],
});
