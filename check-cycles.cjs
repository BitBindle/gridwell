#! /usr/bin/env node

/**
 * Fail when the modules under src/ import each other in a cycle at run time. `npm run lint` runs it.
 *
 * Type-only imports (`import type`, `export type`, or names that are all marked `type`) are not counted: the build
 * removes them. Modules import each other with the .js extension, which stands for the .ts file.
 */
const fs = require('fs');
const path = require('path');
const ts = require('typescript');

const root = path.join(__dirname, 'src');

const listModules = function (dir) {
    let files = [];
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const file = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            files = files.concat(listModules(file));
        } else if (entry.name.endsWith('.ts') && !entry.name.endsWith('.d.ts')) {
            files.push(file);
        }
    }
    return files;
};

// The modules that a module imports at run time
const getImports = function (file) {
    const source = ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.Latest);
    const imports = [];

    for (const statement of source.statements) {
        let specifier = null;

        if (ts.isImportDeclaration(statement)) {
            const clause = statement.importClause;
            const typeOnly =
                clause &&
                (clause.isTypeOnly ||
                    (!clause.name &&
                        clause.namedBindings &&
                        ts.isNamedImports(clause.namedBindings) &&
                        clause.namedBindings.elements.length > 0 &&
                        clause.namedBindings.elements.every((element) => element.isTypeOnly)));
            if (!typeOnly) {
                specifier = statement.moduleSpecifier;
            }
        } else if (ts.isExportDeclaration(statement) && statement.moduleSpecifier && !statement.isTypeOnly) {
            specifier = statement.moduleSpecifier;
        }

        if (specifier && ts.isStringLiteral(specifier) && specifier.text.startsWith('.')) {
            const target = path.resolve(path.dirname(file), specifier.text.replace(/\.js$/, '.ts'));
            if (fs.existsSync(target)) {
                imports.push(target);
            }
        }
    }

    return imports;
};

const graph = new Map();
for (const file of listModules(root)) {
    graph.set(file, getImports(file));
}

// Tarjan's algorithm: each strongly connected component with more than one module, or a module that imports itself, is a cycle
const cycles = [];
const index = new Map();
const low = new Map();
const stack = [];
const onStack = new Set();
let counter = 0;

const visit = function (file) {
    index.set(file, counter);
    low.set(file, counter);
    counter++;
    stack.push(file);
    onStack.add(file);

    for (const target of graph.get(file) || []) {
        if (!index.has(target)) {
            visit(target);
            low.set(file, Math.min(low.get(file), low.get(target)));
        } else if (onStack.has(target)) {
            low.set(file, Math.min(low.get(file), index.get(target)));
        }
    }

    if (low.get(file) === index.get(file)) {
        const component = [];
        let member;
        do {
            member = stack.pop();
            onStack.delete(member);
            component.push(member);
        } while (member !== file);

        if (component.length > 1 || (graph.get(file) || []).includes(file)) {
            cycles.push(component.map((member) => path.relative(root, member)).sort());
        }
    }
};

for (const file of graph.keys()) {
    if (!index.has(file)) {
        visit(file);
    }
}

if (cycles.length) {
    console.error('Import cycles between the modules in src/ (type-only imports are not counted):');
    for (const cycle of cycles) {
        console.error('  ' + cycle.join(', '));
    }
    console.error('A module imports another module of the same cycle at run time. See "No import cycles" in ARCHITECTURE.md.');
    process.exit(1);
}

console.log('No import cycles in src/ (' + graph.size + ' modules)');
