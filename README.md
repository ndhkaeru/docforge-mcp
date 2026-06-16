# DocForge MCP

![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)
![MCP](https://img.shields.io/badge/MCP-stdio-blue)
![Platform](https://img.shields.io/badge/build-Windows%20x64-lightgrey)

> **Local-first MCP servers for converting, reading, and safely editing common document formats.**

DocForge MCP is a collection of focused [Model Context Protocol](https://modelcontextprotocol.io) servers. Each server runs over `stdio` and gives agents format-specific tools for Markdown, Excel, PDF, DOCX, PPTX, CSV, HTML, plain text, and JSON/JSONL.

The project is intentionally split into small binaries so MCP clients can enable only the document capabilities they need.

---

## Highlights

- **Markdown-native editing**: outline, sections, tables, diagrams, code blocks, links, frontmatter, TOC, split/merge, and stats.
- **Excel round-trip editing**: load `.xlsx` workbooks into sessions, mutate sheets/rows/cells/styles, and save safely.
- **Conversion servers**: convert PDF, DOCX, PPTX, CSV, HTML, JSON, and text into Markdown/plain text for agents.
- **Local-first**: no hosted service or telemetry; files stay on the machine running the MCP server.
- **Windows release flow**: GitHub Actions can build standalone `.exe` artifacts with PyInstaller.

---

## Completed MCP Servers

| Server | Purpose | Input |
| --- | --- | --- |
| `md-tools` | Markdown structure editor for headings, sections, tables, diagrams, code blocks, links, TOC, frontmatter, split/merge, and lightweight HTML export. | .md, .markdown |
| `excel-tools` | Round-trip Excel workbook editor with session-based reads, sheet operations, rows/columns, styles, merges, data validation, images, and save/reload support. | .xlsx only for safe editing; macro/binary formats are rejected to avoid data loss. |
| `pdf-tools` | Best-effort PDF-to-Markdown text extraction for search, summarization, and downstream editing. | .pdf |
| `docx-tools` | DOCX-to-Markdown conversion for readable document text, lists, tables, and document structure. | .docx |
| `pptx-tools` | PPTX-to-Markdown conversion that emits slide content in presentation order. | .pptx |
| `csv-tools` | CSV-to-Markdown table converter for quick rendering and insertion into docs. | .csv |
| `html-tools` | HTML-to-Markdown converter for turning markup into agent-friendly text. | .html, .htm |
| `text-tools` | Plain text and Markdown reader exposed through the shared conversion tool name. | .txt, .md |
| `json-tools` | JSON/JSONL plain-text reader for inspection and handoff to agents. | .json, .jsonl |

---

## Quick Start

### 1. Install dependencies for development

```powershell
python -m venv .venv-build
.\.venv-build\Scripts\python.exe -m pip install -U pip
.\.venv-build\Scripts\python.exe -m pip install pytest pyinstaller "mcp[cli]" openpyxl pillow pdfminer.six pdfplumber mammoth python-pptx markdownify beautifulsoup4 lxml defusedxml charset-normalizer
```

### 2. Run tests

```powershell
.\.venv-build\Scripts\python.exe -m pytest -q
```

### 3. Build servers

For local builds, prepare the complete source tree, then use your local `build.ps1` helper or the PyInstaller commands mirrored in `.github/workflows/release.yml`.

Build outputs are written to `dist\*-tools.exe` and are intentionally ignored by git.

---

## Client Configuration

Register the desired executable as a `stdio` MCP server. Example for `md-tools` on Windows:

```json
{
  "mcpServers": {
    "docforge-md": {
      "command": "E:\\Project\\mcp-server\\DocumentTools\\dist\\md-tools.exe",
      "args": []
    }
  }
}
```

Use one entry per server binary you want enabled, such as `excel-tools.exe`, `pdf-tools.exe`, or `docx-tools.exe`.

---

## Server Documentation

- [`servers/md/README.md`](servers/md/README.md)
- [`servers/excel/README.md`](servers/excel/README.md)
- [`servers/pdf/README.md`](servers/pdf/README.md)
- [`servers/docx/README.md`](servers/docx/README.md)
- [`servers/pptx/README.md`](servers/pptx/README.md)
- [`servers/csv/README.md`](servers/csv/README.md)
- [`servers/html/README.md`](servers/html/README.md)
- [`servers/text/README.md`](servers/text/README.md)
- [`servers/json/README.md`](servers/json/README.md)

---

## CI and Release

- Push/PR workflow: creates `.venv-build`, installs dependencies, compiles `md-tools`, and runs the test suite.
- Release workflow: runs on tags matching `v*.*.*`, prepares the complete source tree, verifies tests, builds all completed Windows `.exe` servers, generates checksums, and creates a GitHub Release.

---

## Repository Hygiene

Only MCP source, tests, docs, and workflows should be committed. Generated outputs and local build helpers are ignored.
