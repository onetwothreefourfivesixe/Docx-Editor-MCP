#!/usr/bin/env node
// Applies replacements to a .docx as native Word tracked changes that a
// reviewer can accept or reject in Word.
//
// Usage:
//   node edit_docx.mjs <file.docx> --edits <edits.json> [--author "Name"] [--out <new.docx>]
//   node edit_docx.mjs <file.docx> --find "old text" --replace "new text" [--author "Name"] [--out <new.docx>]
//
// edits.json holds a list of edits, applied in order:
//   [ { "find": "old text", "replace": "new text" }, { "find": "remove me", "replace": "" } ]
//
// Each edit replaces EVERY occurrence of "find". An empty "replace" makes the
// change a plain deletion. Without --out the input file is overwritten, and
// it's only written if at least one edit matched.
//
// Prints a JSON summary: { "output": …, "edits": [ { "find": …, "replaced": n }, … ] }

import * as fs from 'node:fs';
import { applyTrackedReplacement, openDocx, saveDocx } from './lib/docx.mjs';

const USAGE =
    'Usage: node edit_docx.mjs <file.docx> (--edits <edits.json> | --find <text> --replace <text>) [--author <name>] [--out <file.docx>]';

function parseArgs(args) {
    const options = { file: undefined, edits: undefined, find: undefined, replace: undefined, author: 'AI Assistant', out: undefined };
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        const value = () => {
            if (i + 1 >= args.length) throw new Error(`${arg} needs a value\n${USAGE}`);
            return args[++i];
        };
        if (arg === '--edits') options.edits = value();
        else if (arg === '--find') options.find = value();
        else if (arg === '--replace') options.replace = value();
        else if (arg === '--author') options.author = value();
        else if (arg === '--out') options.out = value();
        else if (!arg.startsWith('--') && options.file === undefined) options.file = arg;
        else throw new Error(`Unexpected argument ${arg}\n${USAGE}`);
    }
    if (!options.file) throw new Error(USAGE);
    return options;
}

function loadEdits(options) {
    if (options.edits !== undefined) {
        // Strip a byte order mark: some editors and PowerShell add one.
        const raw = fs.readFileSync(options.edits, 'utf8').replace(/^﻿/, '');
        const edits = JSON.parse(raw);
        if (!Array.isArray(edits)) throw new Error('The edits file must contain a JSON array');
        for (const [index, edit] of edits.entries()) {
            if (typeof edit?.find !== 'string' || typeof edit?.replace !== 'string') {
                throw new Error(`Edit ${index + 1} needs string "find" and "replace" fields`);
            }
        }
        return edits;
    }
    if (options.find === undefined || options.replace === undefined) throw new Error(USAGE);
    return [{ find: options.find, replace: options.replace }];
}

function main(args) {
    const options = parseArgs(args);
    const edits = loadEdits(options);
    const opened = openDocx(options.file);
    const date = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');

    const results = edits.map((edit) => ({
        find: edit.find,
        replaced: applyTrackedReplacement(opened, { ...edit, author: options.author, date }),
    }));

    const changed = results.some((result) => result.replaced > 0);
    const output = options.out ?? options.file;
    if (changed) saveDocx(opened, output);

    console.log(JSON.stringify({ output: changed ? output : null, edits: results }, null, 2));
    if (!changed) console.error('No edits matched, so no file was written.');
}

try {
    main(process.argv.slice(2));
} catch (error) {
    console.error(`Error: ${error.message}`);
    process.exit(1);
}
