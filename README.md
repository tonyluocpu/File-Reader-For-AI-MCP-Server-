# AI File Bridge

A Model Context Protocol (MCP) server that enables AI assistants to read and extract content from file formats they cannot handle directly. Currently supports Microsoft Word documents (.docx), with plans to expand to PDF, Excel, PowerPoint, and other popular formats that AI assistants cannot process natively.

[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue?style=flat&logo=github)](https://github.com/tonyluocpu/ai-file-bridge)
[![Python](https://img.shields.io/badge/Python-3.8+-green?style=flat&logo=python)](https://python.org)
[![MCP](https://img.shields.io/badge/MCP-Compatible-purple?style=flat)](https://modelcontextprotocol.io)

## Why This Tool Exists

**AI assistants (LLMs) cannot directly read many popular file formats.** They can only process plain text, markdown, and a few other formats. This MCP server bridges that gap by:

- **Converting various file formats to readable text** that AI can understand
- **Preserving document structure** (paragraphs, tables, formatting, etc.)
- **Enabling AI analysis** of documents for writing assistance, content review, and more
- **Expanding AI capabilities** to handle files that were previously inaccessible

### Current Support
- ✅ **Microsoft Word** (.docx) - Full text extraction with formatting preservation

### Planned Support
- 🔄 **PDF Documents** (.pdf) - Text and structure extraction
- 🔄 **Microsoft Excel** (.xlsx, .xls) - Spreadsheet data and formulas
- 🔄 **Microsoft PowerPoint** (.pptx, .ppt) - Slide content and notes
- 🔄 **Rich Text Format** (.rtf) - Formatted text documents
- 🔄 **OpenDocument** (.odt, .ods, .odp) - Open source office formats
- 🔄 **E-books** (.epub, .mobi) - Book content extraction

## Project Status

✅ **Production Ready** - Word document support fully tested and verified  
✅ **MCP Compatible** - Works with Claude, Cursor, and other MCP-compatible AI assistants  
✅ **No API Keys Required** - Runs independently without external dependencies  
✅ **Cross-Platform** - Tested on macOS, Linux, and Windows  
🚀 **Expanding** - Actively developing support for additional file formats

## Features

### Current (Word Documents)
- 📄 **Read Word Documents**: Extract text content from .docx files
- 🎨 **Formatting Preservation**: Maintain basic formatting information (bold, italic, underline)
- 📊 **Table Support**: Extract content from tables within documents
- 🌍 **Multi-language Support**: Works with documents in any language (tested with Chinese, English, and more)

### Universal Features
- 🔧 **MCP Compatible**: Integrates seamlessly with AI assistants like Claude, Cursor, and other MCP-compatible tools
- 🖥️ **Cross-platform**: Works on macOS, Linux, and Windows
- 🚀 **Extensible Architecture**: Built to easily add support for new file formats
- 🔒 **No API Keys Required**: Runs completely independently

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. Clone this repository:
```bash
git clone https://github.com/tonyluocpu/ai-file-bridge.git
cd ai-file-bridge
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Make the server executable:
```bash
chmod +x word_reader_mcp_server.py
```

## Usage

### Standalone Usage

You can test the server directly:

```bash
python3 word_reader_mcp_server.py
```

The server will run in stdio mode and wait for JSON-RPC requests.

### Integration with AI Assistants

Once configured, you can ask your AI assistant to:

- **"Read this Word document and summarize it"**
- **"Analyze the writing style in this document"**
- **"Extract all the tables from this Word file"**
- **"Help me improve the grammar in this document"**
- **"Convert this Word document to markdown format"**
- **"What are the main topics discussed in this document?"**
- **"Extract all the key points from this document"**

#### Real-World Use Cases

- **Content Analysis**: Analyze writing styles, themes, and structure
- **Document Summarization**: Create concise summaries of long documents  
- **Data Extraction**: Pull specific information from structured documents
- **Writing Assistance**: Review and improve document content
- **Research**: Process academic papers and reports
- **Translation**: Extract content for translation workflows

#### Cursor IDE

Add the following to your Cursor MCP configuration:

```json
{
  "mcpServers": {
    "word-reader": {
      "command": "python3",
      "args": ["/path/to/word-reader-mcp-server/word_reader_mcp_server.py"],
      "cwd": "/path/to/word-reader-mcp-server"
    }
  }
}
```

#### Claude Desktop

Add to your Claude Desktop MCP configuration:

```json
{
  "mcpServers": {
    "word-reader": {
      "command": "python3",
      "args": ["/path/to/word-reader-mcp-server/word_reader_mcp_server.py"]
    }
  }
}
```

#### Other MCP-Compatible AI Assistants

This server works with any AI assistant that supports the Model Context Protocol (MCP). Check your AI assistant's documentation for MCP configuration instructions.

## Available Tools

### 1. `read_word_document`

Reads and extracts content from a Microsoft Word document.

**Parameters:**
- `file_path` (string, required): Path to the .docx file to read
- `include_formatting` (boolean, optional): Whether to include basic formatting information (default: false)

**Returns:**
- Document content with paragraphs and tables
- Basic formatting information (if requested)
- Total paragraph and table counts

**Example:**
```json
{
  "name": "read_word_document",
  "arguments": {
    "file_path": "/path/to/document.docx",
    "include_formatting": true
  }
}
```

### 2. `list_word_documents`

Lists all .docx files in a specified directory.

**Parameters:**
- `directory_path` (string, optional): Path to the directory to search (default: current directory)

**Returns:**
- List of .docx files with metadata (name, path, size, modification time)

**Example:**
```json
{
  "name": "list_word_documents",
  "arguments": {
    "directory_path": "/path/to/documents"
  }
}
```

## Example Use Cases

### Writing Style Analysis
Perfect for analyzing fiction writing styles, character development patterns, and narrative techniques:

```python
# Extract content and analyze writing patterns
document_content = read_word_document({
    "file_path": "sample_fiction.docx",
    "include_formatting": True
})
```

### Document Processing
Extract and process content from Word documents for various AI tasks:

- Content summarization
- Style transfer
- Character analysis
- Plot structure analysis
- Language learning materials

### Batch Processing
Process multiple documents at once:

```python
# List all documents in a folder
documents = list_word_documents({"directory_path": "/path/to/novels"})

# Process each document
for doc in documents['docx_files']:
    content = read_word_document({"file_path": doc['path']})
    # Process content...
```

## Technical Details

### Architecture
- **Protocol**: JSON-RPC 2.0 over stdio
- **Document Processing**: Uses `python-docx` library
- **Error Handling**: Comprehensive error reporting
- **Async Support**: Full async/await support for MCP communication

### Supported Formats
- Microsoft Word 2007+ (.docx files)
- Legacy Word formats (.doc) are not supported

### Performance
- Efficient memory usage for large documents
- Fast processing of typical document sizes
- Streaming support for very large files

## Development

### Project Structure
```
word-reader-mcp-server/
├── word_reader_mcp_server.py    # Main MCP server
├── requirements.txt             # Python dependencies
├── setup.py                    # Package setup
├── README.md                   # This file
├── LICENSE                     # License file
├── test_mcp_server.py         # Test script
└── examples/                   # Example usage
    └── sample_usage.py
```

### Running Tests
```bash
python3 test_mcp_server.py
```
### Common Issues

**ImportError: No module named 'docx'**
```bash
pip install python-docx
```

**Permission denied when running the server**
```bash
chmod +x word_reader_mcp_server.py
```

**MCP server not connecting**
- Verify the file path in your MCP configuration
- Ensure Python 3.8+ is installed
- Check that all dependencies are installed

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Repository Information

- **GitHub Repository**: [tonyluocpu/ai-file-bridge](https://github.com/tonyluocpu/ai-file-bridge)
- **Clone URL**: `git clone https://github.com/tonyluocpu/ai-file-bridge.git`
- **Issues**: [Report bugs or request features](https://github.com/tonyluocpu/ai-file-bridge/issues)
- **Stars**: ⭐ Star this repository if you find it useful!

## Acknowledgments

- Built using the [python-docx](https://python-docx.readthedocs.io/) library
- Compatible with the [Model Context Protocol](https://modelcontextprotocol.io/) specification
- Inspired by the need for better document processing in AI workflows
- Tested and verified with real-world documents and AI assistants



