// Runs the tests in TypeScript. tsx would use the paths of tsconfig.json, which map jsuites to its typings for the type
// check, also when the bundle requires jsuites, so it gets a tsconfig without them.
process.env.TSX_TSCONFIG_PATH = require('path').join(__dirname, 'tsconfig.run.json');

require('tsx/cjs');
