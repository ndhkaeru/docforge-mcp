"""html-tools — MCP server: convert HTML to Markdown."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('html-tools', instructions='Convert HTML files to Markdown text. Use convert_to_markdown for .html/.htm input when agents need readable Markdown instead of raw markup.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Convert an HTML file to Markdown text."""
    from converter import convert_html
    return convert_html(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
