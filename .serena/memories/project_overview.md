# Enhanced Word MCP Server - Project Overview

## Purpose
This is an Enhanced Word MCP (Model Context Protocol) server for academic research collaboration. It provides advanced tools for Microsoft Word document manipulation including:

- Advanced search and replace with formatting
- Review tools (comments, track changes)
- Section management and organization
- Document protection and signatures
- Footnote/endnote management
- PDF conversion
- Academic paper formatting

## Tech Stack
- **Primary Language**: Python 3.11+
- **Framework**: FastMCP for MCP server implementation
- **Key Dependencies**:
  - python-docx (Word document manipulation)
  - mcp[cli] (Model Context Protocol)
  - msoffcrypto-tool (document encryption)
  - docx2pdf (PDF conversion)
- **Packaging**: Both Python (pyproject.toml) and NPM (package.json) for distribution
- **Entry Point**: Node.js wrapper that spawns Python process

## Architecture
- Main server: `word_document_server/main.py`
- Tools organized by category in `word_document_server/tools/`
- Core utilities in `word_document_server/core/` and `word_document_server/utils/`
- NPM wrapper in `index.js` and `bin/enhanced-word-mcp-server.js`

## Distribution
- Published as NPM package for easy MCP server installation
- Python backend handles actual Word document operations
- Node.js frontend provides cross-platform executable entry point