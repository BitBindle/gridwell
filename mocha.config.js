#! /usr/bin/env node

const { JSDOM } = require('jsdom');

const { window } = new JSDOM('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>', { url: 'https://localhost' });

// Node defines these too, but the jsdom elements accept only the events and the blobs of jsdom. The timers stay those of Node.
const browserGlobals = ['DOMException', 'Event', 'CustomEvent', 'MessageEvent', 'EventTarget', 'Blob', 'File', 'FormData', 'URL', 'atob', 'btoa', 'navigator'];

// Make the window the global object of the tests, as in a browser
for (const key of Object.getOwnPropertyNames(window)) {
    if (!key.startsWith('_') && (!(key in global) || browserGlobals.includes(key))) {
        global[key] = window[key];
    }
}

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
