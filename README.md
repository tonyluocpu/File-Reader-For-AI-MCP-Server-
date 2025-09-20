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

The AI File Bridge provides 5 powerful tools for AI assistants:

### 🔧 Core Reading Tools

#### `read_document`
**Universal document reader** - Automatically detects file type and reads accordingly
- **Input**: `file_path` (required), `include_formatting`, `include_formulas`, `include_formatting_excel`, `include_data_validation`
- **Output**: Structured data with text, tables, formatting, and metadata
- **Use case**: When you don't know the file type or want automatic detection

#### `read_word_document` 
**Microsoft Word document reader** - Extracts text, formatting, and structure from .docx files
- **Input**: `file_path` (required), `include_formatting` (optional)
- **Output**: 
  - Full document text with paragraph structure
  - Table data with formatting
  - Document metadata and properties
- **Use case**: Reading reports, essays, documentation, or any Word document

#### `read_excel_file`
**Microsoft Excel file reader** - Comprehensive Excel data extraction with formulas and formatting
- **Input**: `file_path` (required), `include_formulas`, `include_formatting`, `include_data_validation`
- **Output**:
  - Cell data with values and formulas
  - Multiple worksheet support
  - Cell formatting (colors, fonts, borders)
  - Data validation rules and dropdown lists
  - Chart and graph information
- **Use case**: Analyzing spreadsheets, financial data, reports, or any Excel workbook

### 🔍 Utility Tools

#### `list_supported_files`
**Directory scanner** - Finds all supported file formats in a directory
- **Input**: `directory_path` (optional, defaults to current directory)
- **Output**: List of supported files with metadata (name, size, type, modification date)
- **Use case**: Discovering what files can be processed in a folder

#### `get_supported_formats`
**Format information** - Lists all currently supported file formats
- **Input**: None required
- **Output**: Dictionary of supported file extensions and their descriptions
- **Use case**: Checking what file types the server can handle

### 📊 Example Output Structure

**Word Document Output:**
```json
{
  "success": true,
  "file_path": "document.docx",
  "paragraphs": [...],
  "tables": [...],
  "metadata": {...}
}
```

**Excel File Output:**
```json
{
  "success": true,
  "file_path": "workbook.xlsx",
  "worksheets": [
    {
      "name": "Sheet1",
      "data": [...],
      "formulas": [...],
      "formatting": [...],
      "validation": [...]
    }
  ]
}
```

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



