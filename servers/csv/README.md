# csv-tools

CSV-to-Markdown table converter for quick rendering and insertion into docs.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.csv

## Recommended Workflow

- Call `convert_to_markdown` with the CSV path.
- The response is a Markdown pipe table.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert a CSV file to a Markdown table. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
