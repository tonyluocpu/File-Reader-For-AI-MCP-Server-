# AI File Bridge

A Model Context Protocol (MCP) server that enables AI assistants to read and extract content from file formats they cannot handle directly. Currently supports Microsoft Word documents (.docx) and Excel files (.xlsx, .xls, .xlsm).

[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue?style=flat&logo=github)](https://github.com/tonyluocpu/ai-file-bridge)
[![Python](https://img.shields.io/badge/Python-3.8+-green?style=flat&logo=python)](https://python.org)
[![MCP](https://img.shields.io/badge/MCP-Compatible-purple?style=flat)](https://modelcontextprotocol.io)

## Why This Tool Exists

**AI assistants cannot directly read Excel and Word files.** They can only process plain text and markdown. This MCP server bridges that gap by converting binary file formats into readable text that AI can understand.

## Supported Formats

### ✅ Microsoft Word (.docx)
- Full text extraction with formatting preservation
- Table support and multi-language documents

### ✅ Microsoft Excel (.xlsx, .xls, .xlsm)
- **Formulas & Calculations** - Extract all Excel formulas and results
- **Cell Formatting** - Colors, borders, fonts, number formats
- **Multiple Worksheets** - Read all sheets in a workbook
- **Data Validation** - Dropdown lists, constraints, validation rules
- **Table Structure** - Preserve row/column relationships

## Features

- 🔧 **MCP Compatible** - Works with Cursor, Claude, and other AI assistants
- 🖥️ **Cross-platform** - macOS, Linux, Windows
- 🚀 **No API Keys** - Runs independently
- 📊 **Comprehensive Data** - Extracts formulas, formatting, and structure

## Installation

1. Clone this repository:
```bash
git clone https://github.com/tonyluocpu/ai-file-bridge.git
cd ai-file-bridge
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Cursor IDE Setup

Add to your Cursor MCP configuration:

```json
{
  "mcpServers": {
    "ai-file-bridge": {
      "command": "python3",
      "args": ["/path/to/ai-file-bridge/ai_file_bridge_server.py"],
      "cwd": "/path/to/ai-file-bridge"
    }
  }
}
```

### Example AI Assistant Commands

Once configured, ask your AI assistant:

- **"Read this Excel file and analyze the data"**
- **"What formulas are used in this spreadsheet?"**
- **"Extract all the data from this Excel workbook"**
- **"Read this Word document and summarize it"**
- **"Analyze the writing style in this document"**
- **"What validation rules are applied to this Excel file?"**

## Available Tools

- `read_document` - Universal document reader for Word and Excel files
- `read_word_document` - Extract text and formatting from Word documents
- `read_excel_file` - Extract data, formulas, and formatting from Excel files
- `list_supported_files` - Find all supported files in a directory
- `get_supported_formats` - List all supported file formats

## Testing

Run the test suite to verify everything works:

```bash
python3 test_excel_functionality.py
python3 test_mcp_connection.py
```

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Repository

- **GitHub**: [tonyluocpu/ai-file-bridge](https://github.com/tonyluocpu/ai-file-bridge)
- **Clone**: `git clone https://github.com/tonyluocpu/ai-file-bridge.git`



