---
name: docx-tracked-edits
description: Propose edits to Word .docx files as native tracked changes that a reviewer can accept or reject in Microsoft Word. Use when the user asks to "redline this document", "suggest edits in track changes", "mark up my .docx", "copy-edit this Word file", or "propose changes to the contract", or wants edits credited to a named author instead of silently rewriting the file.
---

# Tracked edits in Word documents

This skill edits `.docx` files without rewriting them silently. Every replacement becomes a native Word tracked change: a deletion and an insertion credited to an author, which the user accepts or rejects in Word (**Review → Track Changes**).

It uses two Node.js scripts in this skill's `scripts/` folder. They need Node.js 18 or later and nothing else. There are no packages to install, they make no network calls, and they only read and write the files you name.

| Script | What it does |
| --- | --- |
| `scripts/read_docx.mjs` | Prints the document text, one numbered paragraph per line. Read-only. |
| `scripts/edit_docx.mjs` | Applies replacements as tracked changes. |

Below, `<skill>` means this skill's folder. Use the real path when running the commands.

## Workflow

### 1. Read the document first

```sh
node <skill>/scripts/read_docx.mjs "/path/to/file.docx"
```

The output shows the exact text you can match. Never guess wording. Copy it from this output. Markers in the output:

- `[|]` is a boundary a match can't cross: a tab, line break, field, or hyperlink edge.
- `{+…+}` and `{-…-}` are existing tracked insertions and deletions. Their text can't be edited until someone accepts or rejects them.

For long documents, add `--find "some text"` to print only the paragraphs containing that text, plus the number of editable matches.

### 2. Plan the edits

- **Each edit replaces every occurrence** of its `find` text. Include enough surrounding words that the `find` text only matches where you intend. Check with `read_docx.mjs --find "…"`: the match count it reports is the number the edit will replace.
- **Keep edits small.** Target the phrase that changes, not a whole paragraph. Small redlines are easier to review.
- **Stay within one paragraph** and don't cross a `[|]` marker. Split such changes into separate edits.
- **Copy typography exactly**: curly quotes (`’ “ ”`), dashes (`– —`), and non-breaking spaces must match the document.
- An empty `replace` makes the change a plain deletion.
- The inserted text takes the formatting at the start of the match (for example, bold).

### 3. Apply the edits

Write the edits to a JSON file. This avoids shell-quoting problems with quotes and special characters:

```json
[
  { "find": "revenue grew by twelve percent", "replace": "revenue grew by fifteen percent" },
  { "find": "We will utilize", "replace": "We will use" },
  { "find": " (see appendix)", "replace": "" }
]
```

Then run:

```sh
node <skill>/scripts/edit_docx.mjs "/path/to/file.docx" --edits edits.json --author "Name" --out "/path/to/file.edited.docx"
```

- `--author` sets the name shown on each revision (default `AI Assistant`). Use the name the user gives you.
- `--out` writes a new file. **Use it by default**, so the original stays untouched. Leave it out only if the user explicitly asks to edit the file in place, because without it the input file is overwritten.
- For a single edit, `--find "old" --replace "new"` can be used instead of `--edits`.
- Edits are applied in order, and text already inside a tracked change (including one made by an earlier edit in the same run) isn't matched again.

The script prints a JSON summary with the number of replacements per edit. It writes nothing if no edit matched.

### 4. Check and report

- **`replaced` higher than expected**: an unintended match changed too. Tell the user, since they can reject that revision in Word, or redo the run with a more specific `find` from the original file.
- **`replaced: 0`**: the text differs slightly (quotes, spacing, a `[|]` boundary, or it's inside an existing tracked change). Re-read with `--find` and retry with exact text. Don't widen the match until it hits something you didn't intend.
- **Error**: the file is missing, isn't a valid `.docx`, is encrypted, or is locked because it's open in Word. If it's locked, ask the user to close it.
- Summarize the proposed changes for the user and tell them to review them in Word under **Review → Track Changes**.

## Limits

- Only the main body is edited: paragraphs and tables, including text boxes in the body. Headers, footers, footnotes, endnotes, and comments are not.
- It replaces text only. It doesn't add comments, make formatting-only changes, or accept or reject existing revisions.
- `.doc` (legacy binary), `.odt`, password-protected, and ZIP64 (over 4 GB) files aren't supported. For `.doc` or `.odt`, ask the user to save a copy as `.docx` first.
