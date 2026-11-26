#!/usr/bin/env node
// Prints the main-body text of a .docx, one numbered paragraph per line, so
// you can copy exact text into an edit. Read-only: the file is not changed.
//
// Usage:
//   node read_docx.mjs <file.docx> [--find "text"]
//
// Markers in the output (they're not part of the document text):
//   [|]      a boundary a match can't cross (tab, line break, field, hyperlink edge…)
//   {+…+}    text inserted by an existing tracked change (can't be edited)
//   {-…-}    text deleted by an existing tracked change (can't be edited)
//
// --find prints only the paragraphs containing the text, plus the number of
// editable matches, which is the number an edit would replace.

import { BARRIER, describeParagraph, indexParagraph, openDocx, paragraphs } from './lib/docx.mjs';

function main(args) {
    const filePath = args[0];
    const findIndex = args.indexOf('--find');
    const find = findIndex >= 0 ? args[findIndex + 1] : undefined;
    if (!filePath || filePath.startsWith('--') || (findIndex >= 0 && !find)) {
        console.error('Usage: node read_docx.mjs <file.docx> [--find "text"]');
        process.exit(2);
    }

    const { doc } = openDocx(filePath);
    let total = 0;
    paragraphs(doc).forEach((paragraph, index) => {
        const shown = describeParagraph(paragraph);
        if (!shown) return;
        if (find !== undefined) {
            const { text } = indexParagraph(paragraph);
            let count = 0;
            for (let at = text.indexOf(find); at !== -1 && !find.includes(BARRIER); at = text.indexOf(find, at + find.length)) count++;
            if (count === 0) return;
            total += count;
        }
        console.log(`¶${index + 1}: ${shown}`);
    });
    if (find !== undefined) console.log(`\n${total} editable match(es) for ${JSON.stringify(find)}`);
}

try {
    main(process.argv.slice(2));
} catch (error) {
    console.error(`Error: ${error.message}`);
    process.exit(1);
}
