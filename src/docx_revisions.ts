import JSZip from 'jszip';
import { DOMParser, XMLSerializer, type Element, type Node } from '@xmldom/xmldom';
import * as fs from 'fs/promises';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ELEMENT_NODE = 1;

// Marks a point in a paragraph's text that a match must not span.
// U+FFFF is not a legal XML character, so it never occurs in document text.
const BARRIER = '\uFFFF';

// Paragraph children that do not interrupt the text flowing around them.
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

interface ReplaceOptions {
    filePath: string;
    outputPath?: string;
    targetText: string;
    replacementText: string;
    author?: string;
}

interface TextPiece {
    node: Element; // <w:t>
    run: Element; // <w:r> that owns the <w:t>
    start: number;
    end: number;
}

interface Match {
    start: number;
    end: number;
    runProperties: Element | undefined;
}

function wordName(node: Node): string | null {
    if (node.nodeType !== ELEMENT_NODE || node.namespaceURI !== W_NS) return null;
    return node.localName;
}

function getRunProperties(run: Element): Element | undefined {
    return Array.from(run.childNodes).find((child) => wordName(child) === 'rPr') as Element | undefined;
}

/**
 * Flattens a paragraph into one string plus the <w:t> pieces it came from.
 * Text in existing tracked changes and nested paragraphs is not included.
 */
function indexParagraph(paragraph: Element): { text: string; pieces: TextPiece[] } {
    let text = '';
    const pieces: TextPiece[] = [];

    const indexRun = (run: Element) => {
        for (const child of Array.from(run.childNodes)) {
            if (child.nodeType !== ELEMENT_NODE) continue;
            const name = wordName(child);
            if (name === 'rPr' || name === 'lastRenderedPageBreak') continue;
            if (name === 't') {
                const value = child.textContent || '';
                pieces.push({ node: child as Element, run, start: text.length, end: text.length + value.length });
                text += value;
            } else {
                text += BARRIER;
            }
        }
    };

    const walk = (parent: Node) => {
        for (const child of Array.from(parent.childNodes)) {
            if (child.nodeType !== ELEMENT_NODE) continue;
            const name = wordName(child);
            if (name === 'r') {
                indexRun(child as Element);
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

function findMatches(text: string, pieces: TextPiece[], target: string): Match[] {
    const matches: Match[] = [];
    let index = text.indexOf(target);
    while (index !== -1) {
        const start = index;
        const firstPiece = pieces.find((piece) => piece.start <= start && start < piece.end);
        matches.push({
            start,
            end: start + target.length,
            runProperties: firstPiece && getRunProperties(firstPiece.run),
        });
        index = text.indexOf(target, start + target.length);
    }
    return matches;
}

export async function applyTrackedReplacement(opts: ReplaceOptions): Promise<number> {
    if (!opts.targetText) {
        throw new Error('targetText must not be empty');
    }
    if (opts.targetText.includes(BARRIER)) {
        throw new Error('targetText contains an invalid character');
    }

    const fileBuffer = await fs.readFile(opts.filePath);
    const zip = await JSZip.loadAsync(fileBuffer);

    const docXmlEntry = zip.file('word/document.xml');
    if (!docXmlEntry) {
        throw new Error('Invalid .docx file: word/document.xml missing');
    }

    const xmlContent = await docXmlEntry.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlContent, 'application/xml');

    // Find max existing revision ID to prevent collisions
    let currentId = 1;
    const existingRevs = doc.getElementsByTagNameNS(W_NS, 'ins');
    const existingDels = doc.getElementsByTagNameNS(W_NS, 'del');
    const allRevs = Array.from(existingRevs).concat(Array.from(existingDels));
    for (const el of allRevs) {
        const idVal = parseInt(el.getAttribute('w:id') || '0', 10);
        if (!isNaN(idVal) && idVal >= currentId) currentId = idVal + 1;
    }

    const nowIso = new Date().toISOString();
    const author = opts.author || 'AI Assistant';
    let matchCount = 0;

    const createRun = (runProperties: Element | undefined): Element => {
        const run = doc.createElementNS(W_NS, 'w:r');
        if (runProperties) {
            run.appendChild(runProperties.cloneNode(true));
        }
        return run;
    };

    const createText = (tagName: 'w:t' | 'w:delText', value: string): Element => {
        const el = doc.createElementNS(W_NS, tagName);
        el.setAttribute('xml:space', 'preserve');
        el.textContent = value;
        return el;
    };

    const createRevision = (tagName: 'w:ins' | 'w:del', run: Element): Element => {
        const el = doc.createElementNS(W_NS, tagName);
        el.setAttribute('w:id', String(currentId++));
        el.setAttribute('w:author', author);
        el.setAttribute('w:date', nowIso);
        el.appendChild(run);
        return el;
    };

    // Replaces a run with kept runs and tracked revisions, leaving unmatched text and other content untouched.
    const rewriteRun = (run: Element, text: string, pieces: TextPiece[], matches: Match[]) => {
        const parent = run.parentNode;
        if (!parent) return;

        const runProperties = getRunProperties(run);
        let keptRun: Element | null = null;
        const keep = (node: Node) => {
            if (!keptRun) {
                keptRun = createRun(runProperties);
                parent.insertBefore(keptRun, run);
            }
            keptRun.appendChild(node);
        };

        for (const child of Array.from(run.childNodes)) {
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
                if (from > position) {
                    keep(createText('w:t', text.slice(position, from)));
                }
                keptRun = null;

                const delRun = createRun(runProperties);
                delRun.appendChild(createText('w:delText', text.slice(from, to)));
                parent.insertBefore(createRevision('w:del', delRun), run);

                if (to === match.end && opts.replacementText) {
                    const insRun = createRun(match.runProperties);
                    insRun.appendChild(createText('w:t', opts.replacementText));
                    parent.insertBefore(createRevision('w:ins', insRun), run);
                }
                position = to;
            }
            if (position < piece.end) {
                keep(createText('w:t', text.slice(position, piece.end)));
            }
        }

        parent.removeChild(run);
    };

    // Snapshot the live list, since the loop below changes the document.
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, 'p'));

    for (const paragraph of paragraphs) {
        const { text, pieces } = indexParagraph(paragraph);
        const matches = findMatches(text, pieces, opts.targetText);
        if (matches.length === 0) continue;
        matchCount += matches.length;

        const touchedRuns = new Set(
            pieces
                .filter((piece) => matches.some((match) => match.start < piece.end && match.end > piece.start))
                .map((piece) => piece.run)
        );
        for (const run of touchedRuns) {
            rewriteRun(run, text, pieces, matches);
        }
    }

    if (matchCount > 0) {
        const serializer = new XMLSerializer();
        zip.file('word/document.xml', serializer.serializeToString(doc));
        const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
        await fs.writeFile(opts.outputPath || opts.filePath, outBuffer);
    }

    return matchCount;
}
