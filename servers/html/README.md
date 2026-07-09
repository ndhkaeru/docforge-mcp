# html-tools

HTML-to-Markdown converter for turning markup into agent-friendly text.

## Status

Implemented and included in the completed DocLoupe MCP server set.

## Input

.html, .htm

## Recommended Workflow

- Call `convert_to_markdown` with the HTML path.
- Use output when raw HTML would waste context.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Convert an HTML file to Markdown text. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
