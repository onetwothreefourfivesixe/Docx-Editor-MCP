// Packages agensi/docx-tracked-edits/ into build/docx-tracked-edits.zip with the
// layout Agensi expects: one top-level folder named after the skill, with
// SKILL.md directly inside it. Uses the skill's own dependency-free zip writer.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeEntry, writeZip } from '../agensi/docx-tracked-edits/scripts/lib/zip.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const skillName = 'docx-tracked-edits';
const skillDir = path.join(repo, 'agensi', skillName);
const outFile = path.join(repo, 'build', `${skillName}.zip`);
const SKIP = new Set(['.DS_Store', 'Thumbs.db', 'desktop.ini']);

function collect(dir, relative = '') {
    const files = [];
    for (const item of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
        if (SKIP.has(item.name) || item.name.startsWith('.')) continue;
        const rel = relative ? `${relative}/${item.name}` : item.name;
        const full = path.join(dir, item.name);
        if (item.isSymbolicLink()) throw new Error(`Symlinks are not allowed: ${rel}`);
        if (item.isDirectory()) files.push(...collect(full, rel));
        else files.push({ rel, full });
    }
    return files;
}

const files = collect(skillDir);
if (!files.some((file) => file.rel === 'SKILL.md')) throw new Error('SKILL.md is missing');

const entries = files.map((file) => makeEntry(`${skillName}/${file.rel}`, fs.readFileSync(file.full)));
fs.mkdirSync(path.dirname(outFile), { recursive: true });
fs.writeFileSync(outFile, writeZip(entries));

console.log(`Wrote ${path.relative(repo, outFile)} (${files.length} files):`);
for (const file of files) console.log(`  ${skillName}/${file.rel}`);
