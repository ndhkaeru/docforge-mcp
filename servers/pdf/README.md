# pdf-tools

Best-effort PDF-to-Markdown text extraction for search, summarization, and downstream editing.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.pdf

## Recommended Workflow

- Call `convert_to_markdown` with the PDF path.
- Use the Markdown output for review/search rather than pixel-perfect layout.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert a PDF file to Markdown text. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
