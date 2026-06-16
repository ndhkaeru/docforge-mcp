"""pdf-tools — MCP server: convert PDF to Markdown."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('pdf-tools', instructions='Convert PDF files to Markdown text. Use convert_to_markdown for .pdf input; output is best-effort text extraction suitable for search, summarization, and downstream Markdown editing.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Convert a PDF file to Markdown text."""
    from converter import convert_pdf
    return convert_pdf(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
