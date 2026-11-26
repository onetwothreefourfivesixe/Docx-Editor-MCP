import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { applyTrackedReplacement } from './docx_revisions.js';

const server = new McpServer({
    name: 'docx-editor',
    version: '1.0.0',
});

server.tool(
    'propose_document_edit',
    'Finds text in a .docx Word document and applies a native Word track-changes replacement.',
    {
        filePath: z.string().describe('Absolute path to the .docx file'),
        targetText: z.string().describe('The original exact text to be replaced'),
        replacementText: z.string().describe('The proposed text to insert'),
        author: z.string().optional().describe('Author name to show in Word tracked revisions (defaults to "AI Assistant")'),
    },
    async ({ filePath, targetText, replacementText, author }) => {
        try {
            const count = await applyTrackedReplacement({
                filePath,
                targetText,
                replacementText,
                author: author || 'AI Assistant',
            });

            if (count === 0) {
                return {
                    content: [{ type: 'text', text: `Warning: Target text "${targetText}" was not found.` }],
                };
            }

            return {
                content: [
                    {
                        type: 'text',
                        text: `Successfully replaced ${count} occurrence(s) with native Word track changes.`,
                    },
                ],
            };
        } catch (err: any) {
            return {
                isError: true,
                content: [{ type: 'text', text: `Error updating .docx: ${err.message}` }],
            };
        }
    }
);

async function run() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
}

run().catch((error) => {
    console.error('Fatal server error:', error);
    process.exit(1);
});