# json-tools

JSON/JSONL plain-text reader for inspection and handoff to agents.

## Status

Implemented and included in the completed DocForge MCP server set.

## Input

.json, .jsonl

## Recommended Workflow

- Call `convert_to_markdown` with the JSON/JSONL path.
- Output is raw textual JSON content.

## Tools

| Tool | Description |
| --- | --- |
| `convert_to_markdown` | Return the content of a .json or .jsonl file. |


## Notes

- This server exposes one focused `convert_to_markdown` tool.
- Output is text/Markdown intended for agent consumption, not a pixel-perfect representation of the original file.
