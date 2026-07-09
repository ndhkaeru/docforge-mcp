# pptx-tools

PPTX-to-Markdown conversion that emits slide content in presentation order.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.pptx

## Recommended Workflow

- Call `convert_to_markdown` with the PPTX path.
- Review generated slide sections in order.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert a PPTX file to Markdown text. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
