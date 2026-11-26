# docx-tracked-edits

An agent skill that proposes edits to Word `.docx` files as **native tracked changes**. Every replacement shows up in Word as a deletion and an insertion, credited to the author you choose, that you can accept or reject.

## Requirements

Node.js 18 or later. There are no packages to install.

## What's inside

```
docx-tracked-edits/
├── SKILL.md               instructions the agent follows
├── README.md              this file
└── scripts/
    ├── read_docx.mjs      prints the document text (read-only)
    ├── edit_docx.mjs      applies edits as tracked changes
    └── lib/
        ├── docx.mjs       finds text across formatting runs and writes <w:del>/<w:ins> revisions
        ├── xml.mjs        small XML parser and serializer
        └── zip.mjs        small ZIP reader and writer (uses Node's built-in zlib)
```

## Safety

- It uses only Node built-ins (`fs`, `zlib`). It has no dependencies, makes no network access, and reads no environment variables.
- It reads only the `.docx` and edits file you pass, and writes only the output file. The agent is told to write to a new file (`--out`) unless you ask for in-place edits.
- It never runs shell commands or other programs.

## Example

```sh
node scripts/read_docx.mjs report.docx
node scripts/edit_docx.mjs report.docx --find "utilize" --replace "use" --author "Jane Doe" --out report.edited.docx
```
