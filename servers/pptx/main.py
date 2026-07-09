"""pptx-tools — MCP server: convert PPTX to Markdown."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('pptx-tools', instructions='Convert PPTX files to Markdown text. Use convert_to_markdown for .pptx input; slide text is emitted in presentation order for review and summarization.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Convert a PPTX file to Markdown text."""
    from converter import convert_pptx
    return convert_pptx(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
