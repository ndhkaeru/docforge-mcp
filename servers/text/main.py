"""text-tools — MCP server: read plain text / Markdown files."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('text-tools', instructions='Read plain text and Markdown files as raw text. Use convert_to_markdown for .txt or .md input when a client expects the common conversion tool name.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Return the content of a .txt or .md file."""
    from converter import convert_text
    return convert_text(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
