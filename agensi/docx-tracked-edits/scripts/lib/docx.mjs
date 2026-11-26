// Reads the main body of a .docx and writes replacements into it as native
// Word tracked changes (<w:del> + <w:ins>). Uses only Node built-ins.

import * as fs from 'node:fs';
import { ELEMENT, parseXml, serializeXml } from './xml.mjs';
import { entryContent, readZip, setEntryContent, writeZip } from './zip.mjs';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DOCUMENT_PART = 'word/document.xml';

// Marks a point in a paragraph's text that a match must not span.
// U+FFFF is not a legal XML character, so it never occurs in document text.
export const BARRIER = '￿';

// Paragraph children that don't interrupt the text around them.
const TRANSPARENT_ELEMENTS = new Set([
    'pPr',
    'bookmarkStart',
    'bookmarkEnd',
    'proofErr',
    'commentRangeStart',
    'commentRangeEnd',
    'permStart',
    'permEnd',
]);

// Paragraph children that hold runs which can be edited in place.
const RUN_CONTAINERS = new Set(['hyperlink', 'smartTag', 'sdt', 'sdtContent', 'fldSimple', 'customXml']);

// Revision elements whose w:id values a new revision must not reuse.
const REVISION_ELEMENTS = new Set([
    'ins',
    'del',
    'moveFrom',
    'moveTo',
    'rPrChange',
    'pPrChange',
    'sectPrChange',
    'tblPrChange',
    'trPrChange',
    'tcPrChange',
    'tblGridChange',
    'numberingChange',
    'cellIns',
    'cellDel',
    'cellMerge',
]);

/** Local name of a WordprocessingML element, or null for anything else. */
function wordName(node) {
    if (node.type !== ELEMENT || node.ns !== W_NS) return null;
    return node.localName;
}

function getRunProperties(run) {
    return run.children.find((child) => wordName(child) === 'rPr');
}

/** Opens a .docx and parses its main document part. */
export function openDocx(filePath) {
    const entries = readZip(fs.readFileSync(filePath));
    const entry = entries.find((candidate) => candidate.name === DOCUMENT_PART);
    if (!entry) throw new Error(`Not a Word document: ${DOCUMENT_PART} is missing`);
    const doc = parseXml(entryContent(entry).toString('utf8'));
    const prefix = doc.prefixFor(W_NS);
    if (!prefix) throw new Error('Unsupported document: the WordprocessingML namespace has no prefix');
    return { entries, entry, doc, prefix };
}

/** Writes the edited document part back and saves the package. */
export function saveDocx(opened, outputPath) {
    setEntryContent(opened.entry, Buffer.from(serializeXml(opened.doc), 'utf8'));
    fs.writeFileSync(outputPath, writeZip(opened.entries));
}

/** Every paragraph in the main body, in document order (including nested ones). */
export function paragraphs(doc) {
    return doc.root.findAll(W_NS, 'p');
}

/**
 * Flattens a paragraph into one string plus the <w:t> pieces it came from.
 * Tabs, breaks, fields, existing tracked changes and other non-text content
 * become BARRIER characters, so a match can't span them.
 */
export function indexParagraph(paragraph) {
    let text = '';
    const pieces = [];

    const indexRun = (run) => {
        for (const child of run.children) {
            if (child.type !== ELEMENT) continue;
            const name = wordName(child);
            if (name === 'rPr' || name === 'lastRenderedPageBreak') continue;
            if (name === 't') {
                const value = child.textContent;
                pieces.push({ node: child, run, start: text.length, end: text.length + value.length });
                text += value;
            } else {
                text += BARRIER;
            }
        }
    };

    const walk = (parent) => {
        for (const child of parent.children) {
            if (child.type !== ELEMENT) continue;
            const name = wordName(child);
            if (name === 'r') {
                indexRun(child);
            } else if (name && RUN_CONTAINERS.has(name)) {
                text += BARRIER;
                walk(child);
                text += BARRIER;
            } else if (!name || !TRANSPARENT_ELEMENTS.has(name)) {
                text += BARRIER;
            }
        }
    };

    walk(paragraph);
    return { text, pieces };
}

/**
 * Text of a paragraph for a reader: editable text as is, with "{+…+}" and
 * "{-…-}" around existing tracked insertions and deletions, and "[|]" where
 * a match can't cross (tab, break, field, hyperlink edge…).
 */
export function describeParagraph(paragraph) {
    let out = '';
    const barrier = () => {
        if (!out.endsWith('[|]')) out += '[|]';
    };
    const plainText = (node, textName) => {
        let value = '';
        for (const el of node.findAll(W_NS, textName)) value += el.textContent;
        return value;
    };
    const walk = (parent) => {
        for (const child of parent.children) {
            if (child.type !== ELEMENT) continue;
            const name = wordName(child);
            if (name === 'r') {
                for (const part of child.children) {
                    if (part.type !== ELEMENT) continue;
                    const partName = wordName(part);
                    if (partName === 'rPr' || partName === 'lastRenderedPageBreak') continue;
                    if (partName === 't') out += part.textContent;
                    else barrier();
                }
            } else if (name === 'ins' || name === 'moveTo') {
                barrier();
                out += `{+${plainText(child, 't')}+}`;
                barrier();
            } else if (name === 'del' || name === 'moveFrom') {
                barrier();
                out += `{-${plainText(child, 'delText')}-}`;
                barrier();
            } else if (name && RUN_CONTAINERS.has(name)) {
                barrier();
                walk(child);
                barrier();
            } else if (!name || !TRANSPARENT_ELEMENTS.has(name)) {
                barrier();
            }
        }
    };
    walk(paragraph);
    return out.replace(/^\[\|\]|\[\|\]$/g, '');
}

function findMatches(text, pieces, target) {
    const matches = [];
    let index = text.indexOf(target);
    while (index !== -1) {
        const start = index;
        const firstPiece = pieces.find((piece) => piece.start <= start && start < piece.end);
        matches.push({ start, end: start + target.length, runProperties: firstPiece && getRunProperties(firstPiece.run) });
        index = text.indexOf(target, start + target.length);
    }
    return matches;
}

/**
 * Replaces every occurrence of `find` with `replace` as tracked changes.
 * An empty `replace` makes the change a plain deletion.
 * @returns {number} the number of occurrences replaced
 */
export function applyTrackedReplacement(opened, { find, replace, author, date }) {
    if (!find) throw new Error('The text to find must not be empty');
    if (find.includes(BARRIER)) throw new Error('The text to find contains an invalid character');

    const { doc, prefix } = opened;
    const q = (local) => `${prefix}:${local}`;

    // New revision ids start above every existing one.
    let nextId = 1;
    const scan = (node) => {
        for (const child of node.children) {
            if (child.type !== ELEMENT) continue;
            if (child.ns === W_NS && REVISION_ELEMENTS.has(child.localName)) {
                const id = parseInt(child.getAttribute(q('id')) || '0', 10);
                if (!isNaN(id) && id >= nextId) nextId = id + 1;
            }
            scan(child);
        }
    };
    scan(doc.root);

    const createRun = (runProperties) => {
        const run = doc.createElement(q('r'));
        if (runProperties) run.appendChild(runProperties.clone());
        return run;
    };
    const createText = (localName, value) => {
        const el = doc.createElement(q(localName));
        el.setAttribute('xml:space', 'preserve');
        el.textContent = value;
        return el;
    };
    const createRevision = (localName, run) => {
        const el = doc.createElement(q(localName));
        el.setAttribute(q('id'), String(nextId++));
        el.setAttribute(q('author'), author);
        el.setAttribute(q('date'), date);
        el.appendChild(run);
        return el;
    };

    // Replaces a run with kept runs and tracked revisions, leaving unmatched
    // text and other run content (tabs, breaks, drawings…) untouched.
    const rewriteRun = (run, text, pieces, matches) => {
        const parent = run.parent;
        if (!parent) return;

        const runProperties = getRunProperties(run);
        let keptRun = null;
        const keep = (node) => {
            if (!keptRun) {
                keptRun = createRun(runProperties);
                parent.insertBefore(keptRun, run);
            }
            keptRun.appendChild(node);
        };

        for (const child of [...run.children]) {
            if (child === runProperties) continue;

            const piece = pieces.find((candidate) => candidate.node === child);
            if (!piece) {
                keep(child);
                continue;
            }

            let position = piece.start;
            for (const match of matches) {
                if (match.end <= piece.start || match.start >= piece.end) continue;

                const from = Math.max(match.start, piece.start);
                const to = Math.min(match.end, piece.end);
                if (from > position) keep(createText('t', text.slice(position, from)));
                keptRun = null;

                const deletedRun = createRun(runProperties);
                deletedRun.appendChild(createText('delText', text.slice(from, to)));
                parent.insertBefore(createRevision('del', deletedRun), run);

                // The insertion goes right after the last piece of the match.
                if (to === match.end && replace) {
                    const insertedRun = createRun(match.runProperties);
                    insertedRun.appendChild(createText('t', replace));
                    parent.insertBefore(createRevision('ins', insertedRun), run);
                }
                position = to;
            }
            if (position < piece.end) keep(createText('t', text.slice(position, piece.end)));
        }

        parent.removeChild(run);
    };

    let count = 0;
    for (const paragraph of paragraphs(doc)) {
        const { text, pieces } = indexParagraph(paragraph);
        const matches = findMatches(text, pieces, find);
        if (matches.length === 0) continue;
        count += matches.length;

        const touchedRuns = new Set(
            pieces
                .filter((piece) => matches.some((match) => match.start < piece.end && match.end > piece.start))
                .map((piece) => piece.run)
        );
        for (const run of touchedRuns) rewriteRun(run, text, pieces, matches);
    }
    return count;
}
