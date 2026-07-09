# docx-tools

DOCX-to-Markdown conversion for readable document text, lists, tables, and document structure.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.docx

## Recommended Workflow

- Call `convert_to_markdown` with the DOCX path.
- Use output as Markdown text for agents or documentation pipelines.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert a DOCX file to Markdown text. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
