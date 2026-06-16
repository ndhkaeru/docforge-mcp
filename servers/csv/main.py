"""csv-tools — MCP server: convert CSV to Markdown."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP('csv-tools', instructions='Convert CSV files to clean Markdown tables. Use convert_to_markdown with an absolute or workspace-relative .csv path; the response is raw Markdown table text for direct rendering or insertion into docs.')


@mcp.tool()
def convert_to_markdown(file_path: str) -> str:
    """Convert a CSV file to a Markdown table."""
    from converter import convert_csv
    return convert_csv(file_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
