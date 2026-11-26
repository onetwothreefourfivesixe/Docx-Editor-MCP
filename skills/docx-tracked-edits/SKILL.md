---
name: docx-tracked-edits
description: Propose edits to Word .docx files as native tracked changes (redlines) that a reviewer can accept or reject in Microsoft Word. Use when the user wants to suggest, propose, redline, mark up, copy-edit, or revise a .docx document with track changes, or wants edits credited to a named author, rather than silently rewriting the file.
---

# Tracked edits in Word documents

The `propose_document_edit` tool from the `docx-editor` MCP server replaces text in a `.docx` file. Each replacement becomes a native Word tracked change: a deletion plus an insertion, credited to an author, that the reviewer can accept or reject in Word.

## The tool

`propose_document_edit` takes:

- `filePath`: the **absolute** path to the `.docx` file
- `targetText`: the exact text to replace (must not be empty)
- `replacementText`: the new text. An empty string makes the change a plain deletion.
- `author` (optional): the name shown on the revision. Defaults to `AI Assistant`. If the user gives a name or initials, use them.

It replaces **every** occurrence of `targetText` and reports how many it replaced, or warns that it found nothing.

## Before editing

1. **The file is overwritten in place.** If the user hasn't said a backup is unnecessary, offer to copy the file first (for example to `report.original.docx`).
2. **Get the exact current text.** `targetText` must match character for character. If you haven't seen the document text, read it before proposing edits. For example, extract `word/document.xml` from the zip, or use a docx reading tool if you have one. Don't guess at wording.
3. **Watch typography.** Word often stores curly quotes (`’ “ ”`), en and em dashes (`– —`) and non-breaking spaces. Copy these exactly as they appear in the document, not as straight ASCII.

## Writing good edits

- **Make each target unique.** Since every occurrence is replaced, include enough surrounding words that `targetText` only matches where you mean it. Example: to fix one "its" in "its' value", target `its' value`, not `its`.
- **Keep changes small and focused.** Target the phrase or sentence that changes, not a whole paragraph. Small redlines are easier to review. Make one call per independent change.
- **Stay within one paragraph.** A match can't cross a paragraph break, tab, line break, field, or hyperlink boundary. Split edits that span those into separate calls.
- **Formatting:** a match can span mixed formatting (such as a partly bold phrase). The inserted text takes the formatting at the start of the match.

## What the tool can't do

- It searches only the main body. Headers, footers, footnotes, endnotes, comments and text boxes outside the body flow aren't edited. Tell the user if their edit falls in one of those.
- Text that is already inside a tracked change is skipped. If a target sits within a pending revision, tell the user that change needs to be accepted or rejected first.
- It doesn't add comments, change formatting only, or accept/reject existing revisions.

## Handling results

- **"Successfully replaced N occurrence(s)"**: check N against what you expected. If it's higher, an unintended match was also changed. Tell the user, since they can reject the extra revision in Word.
- **"Target text … was not found"**: the text probably differs slightly (quotes, spacing, a line break, or an earlier edit already changed it). Re-read the document and retry with the exact text. Don't loosen the target until it matches something unintended.
- **Error**: the file is missing, isn't a valid `.docx`, or is locked (for example, open in Word on Windows). Report it, and if it's locked, ask the user to close the file.

When you finish, summarize the changes you proposed and remind the user to review them in Word under **Review → Track Changes**.
