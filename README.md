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

## Install as a Claude Code plugin

The repo is both a plugin and a single-plugin marketplace. The plugin bundles the MCP server and a `docx-tracked-edits` skill that teaches Claude how to use the tool safely.

```sh
/plugin marketplace add <github-user>/Docx-Editor-MCP
/plugin install docx-editor@docx-editor-marketplace
```

Only Node.js 18+ is needed. The server ships prebuilt as `dist/index.js` with all dependencies bundled, so there's no `npm install` step.

## Usage with other MCP clients

Add the server to your client's config, for example Claude Desktop's `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "node",
      "args": ["/absolute/path/to/Docx-Editor-MCP/dist/index.js"]
    }
  }
}
```

Or register it with Claude Code directly, without the plugin:

```sh
claude mcp add docx-editor -- node /absolute/path/to/Docx-Editor-MCP/dist/index.js
```

The server talks over stdio.

## Standalone skill (Agensi and other skill marketplaces)

[agensi/docx-tracked-edits/](agensi/docx-tracked-edits/) is a self-contained version for marketplaces that take a zip with a `SKILL.md`, where an MCP server can't be installed. Instead of calling the MCP tool, the agent runs two small Node scripts (`read_docx.mjs` and `edit_docx.mjs`). They have no dependencies and use only Node built-ins for the ZIP and XML handling, so the package has no `node_modules` and no minified code.

```sh
npm run package:skill   # writes build/docx-tracked-edits.zip (one top-level folder with SKILL.md inside)
```

The revision logic in `agensi/docx-tracked-edits/scripts/lib/docx.mjs` is a port of `src/docx_revisions.ts`. If you change one, change the other.

## Development

```sh
npm install
npm run typecheck   # tsc --noEmit
npm run build       # bundles src/ into dist/index.js with esbuild
npm run dev         # runs src/index.ts directly with tsx
```

`dist/index.js` is committed because the plugin runs it. Rebuild and commit it whenever you change `src/`.

## Project layout

- `src/index.ts`: MCP server and tool definition
- `src/docx_revisions.ts`: unzips the `.docx`, finds matches across runs and writes the `<w:del>`/`<w:ins>` revisions
- `dist/index.js`: bundled server that the plugin runs
- `skills/docx-tracked-edits/SKILL.md`: skill that tells Claude when and how to use the tool
- `agensi/docx-tracked-edits/`: standalone skill with dependency-free scripts, for zip-based skill marketplaces
- `tools/package-skill.mjs`: builds the marketplace zip into `build/`
- `.mcp.json`: starts the bundled server when the plugin is enabled
- `.claude-plugin/plugin.json`, `.claude-plugin/marketplace.json`: plugin and marketplace manifests
