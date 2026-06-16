"""json-tools — MCP server: read JSON / JSONL files."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('json-tools', instructions='Read JSON and JSONL files as plain text for agent workflows. Use convert_to_markdown for .json/.jsonl input when structured data should be inspected without specialized tooling.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Return the content of a .json or .jsonl file."""
    from converter import convert_json
    return convert_json(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
