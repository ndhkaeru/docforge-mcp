"""docx-tools — MCP server: convert DOCX to Markdown."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('docx-tools', instructions='Convert DOCX files to Markdown text. Use convert_to_markdown for .docx input; the converter preserves readable document structure, lists, tables, and extracted text where possible.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Convert a DOCX file to Markdown text."""
    from converter import convert_docx
    return convert_docx(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
