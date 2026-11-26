# Docx Editor MCP

Docx Editor MCP is an MCP server that edits Word (`.docx`) documents by proposing changes as **native Word tracked changes**. Each replacement shows up in Word as a deletion and an insertion that a reviewer can accept or reject, credited to the author you choose.

## Tool

### `propose_document_edit`

Finds every occurrence of a piece of text and replaces it as a tracked change.

| Parameter         | Type   | Required | Description                                                      |
| ----------------- | ------ | -------- | ---------------------------------------------------------------- |
| `filePath`        | string | yes      | Absolute path to the `.docx` file                                |
| `targetText`      | string | yes      | The exact text to replace                                        |
| `replacementText` | string | yes      | The new text (an empty string makes the change a plain deletion) |
| `author`          | string | no       | Name shown on the revision in Word. Defaults to `AI Assistant`   |

It returns the number of replaced occurrences, or a warning if the text was not found. It returns an error if the file can't be read as a `.docx` or if `targetText` is empty.

#### Behavior

- The file is **overwritten in place**. Keep a copy if you need the original.
- A match can span several runs, so text with mixed formatting is found. The inserted text takes the formatting of the start of the match.
- A match must sit inside one paragraph. It won't cross paragraph breaks, tabs, line breaks, fields or hyperlink boundaries.
- Text that is already inside a tracked change is ignored.
- Only the main body (`word/document.xml`) is searched. Headers, footers, footnotes, endnotes and comments are not.

## Setup

Requires Node.js 18+.

```sh
npm install
npx tsc          # compiles src/*.ts to src/*.js (plus .d.ts and source maps)
```

## Usage with an MCP client

Add the server to your client's config, for example Claude Desktop's `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "node",
      "args": ["/absolute/path/to/docx_editor_mcp/src/index.js"]
    }
  }
}
```

For Claude Code:

```sh
claude mcp add docx-editor -- node /absolute/path/to/docx_editor_mcp/src/index.js
```

During development you can skip compiling and run the TypeScript directly with `npx tsx src/index.ts`.

The server talks over stdio.

## Project layout

- `src/index.ts`: MCP server and tool definition
- `src/docx_revisions.ts`: unzips the `.docx`, finds matches across runs and writes the `<w:del>`/`<w:ins>` revisions
