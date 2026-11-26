#! /usr/bin/env node

require('jsdom-global')(undefined, { url: 'https://localhost' });

// jsdom-global does not make Option a global. A cell with parseHTML uses it to remove the scripts.
global.Option = window.Option;

global.root = document.createElement('div');
global.root.style.width = '100%';
global.root.style.height = '100%';
global.root.innerHTML = '';
document.body.appendChild(global.root);

exports.mochaHooks = {
    afterEach(done) {
        // destroy datagrid component
        global.root.innerHTML = '';
        done();
    },
};
