# text-tools

Plain text and Markdown reader exposed through the shared conversion tool name.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.txt, .md

## Recommended Workflow

- Call `convert_to_markdown` for raw text content.
- Use this server when arbitrary text reads are more appropriate than structure-aware md-tools.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Return the content of a .txt or .md file. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
