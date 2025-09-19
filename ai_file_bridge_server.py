#!/usr/bin/env python3
"""
AI File Bridge MCP Server
Universal file reader for AI assistants - supports Word, Excel, and more formats
"""

import asyncio
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Import our custom readers
from excel_reader import ExcelReader
from word_reader import WordReader


class AIFileBridgeServer:
    def __init__(self):
        self.server_name = "ai-file-bridge"
        self.version = "2.0.0"
        self.excel_reader = ExcelReader()
        self.word_reader = WordReader()
        self.supported_formats = {
            '.docx': 'Microsoft Word Document',
            '.xlsx': 'Microsoft Excel Workbook',
            '.xls': 'Microsoft Excel Workbook (Legacy)',
            '.xlsm': 'Microsoft Excel Macro-Enabled Workbook'
        }
    
    async def handle_request(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle incoming MCP requests"""
        method = request.get("method")
        params = request.get("params", {})
        
        if method == "initialize":
            return await self._handle_initialize(params)
        elif method == "tools/list":
            return await self._handle_tools_list()
        elif method == "tools/call":
            return await self._handle_tools_call(params)
        else:
            return {
                "jsonrpc": "2.0",
                "id": request.get("id"),
                "error": {
                    "code": -32601,
                    "message": f"Method not found: {method}"
                }
            }
    
    async def _handle_initialize(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle initialization request"""
        return {
            "jsonrpc": "2.0",
            "id": params.get("id"),
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {}
                },
                "serverInfo": {
                    "name": self.server_name,
                    "version": self.version
                }
            }
        }
    
    async def _handle_tools_list(self) -> Dict[str, Any]:
        """Return list of available tools"""
        return {
            "jsonrpc": "2.0",
            "result": {
                "tools": [
                    {
                        "name": "read_document",
                        "description": "Read and extract content from supported document formats (Word, Excel, etc.)",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "file_path": {
                                    "type": "string",
                                    "description": "Path to the document file to read"
                                },
                                "include_formatting": {
                                    "type": "boolean",
                                    "description": "Whether to include formatting information",
                                    "default": True
                                },
                                "include_formulas": {
                                    "type": "boolean", 
                                    "description": "Whether to include formulas (Excel files)",
                                    "default": True
                                },
                                "include_validation": {
                                    "type": "boolean",
                                    "description": "Whether to include data validation rules (Excel files)",
                                    "default": True
                                }
                            },
                            "required": ["file_path"]
                        }
                    },
                    {
                        "name": "read_excel_file",
                        "description": "Read and extract comprehensive data from Excel files (.xlsx, .xls, .xlsm)",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "file_path": {
                                    "type": "string",
                                    "description": "Path to the Excel file to read"
                                },
                                "include_formatting": {
                                    "type": "boolean",
                                    "description": "Whether to include cell formatting information",
                                    "default": True
                                },
                                "include_formulas": {
                                    "type": "boolean",
                                    "description": "Whether to include formulas and calculations",
                                    "default": True
                                },
                                "include_validation": {
                                    "type": "boolean",
                                    "description": "Whether to include data validation rules",
                                    "default": True
                                }
                            },
                            "required": ["file_path"]
                        }
                    },
                    {
                        "name": "read_word_document",
                        "description": "Read and extract text content from Microsoft Word documents (.docx files)",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "file_path": {
                                    "type": "string",
                                    "description": "Path to the .docx file to read"
                                },
                                "include_formatting": {
                                    "type": "boolean",
                                    "description": "Whether to include basic formatting information",
                                    "default": False
                                }
                            },
                            "required": ["file_path"]
                        }
                    },
                    {
                        "name": "list_supported_files",
                        "description": "List all supported document files in a directory",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "directory_path": {
                                    "type": "string",
                                    "description": "Path to the directory to search for supported files",
                                    "default": "."
                                }
                            }
                        }
                    },
                    {
                        "name": "get_supported_formats",
                        "description": "Get list of all supported file formats",
                        "inputSchema": {
                            "type": "object",
                            "properties": {}
                        }
                    }
                ]
            }
        }
    
    async def _handle_tools_call(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle tool call requests"""
        tool_name = params.get("name")
        arguments = params.get("arguments", {})
        
        try:
            if tool_name == "read_document":
                result = await self._read_document(arguments)
            elif tool_name == "read_excel_file":
                result = await self._read_excel_file(arguments)
            elif tool_name == "read_word_document":
                result = await self._read_word_document(arguments)
            elif tool_name == "list_supported_files":
                result = await self._list_supported_files(arguments)
            elif tool_name == "get_supported_formats":
                result = await self._get_supported_formats(arguments)
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": params.get("id"),
                    "error": {
                        "code": -32601,
                        "message": f"Tool not found: {tool_name}"
                    }
                }
            
            return {
                "jsonrpc": "2.0",
                "id": params.get("id"),
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result, indent=2, ensure_ascii=False, default=str)
                        }
                    ]
                }
            }
        except Exception as e:
            return {
                "jsonrpc": "2.0",
                "id": params.get("id"),
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {str(e)}"
                }
            }
    
    async def _read_document(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Universal document reader - automatically detects file type"""
        file_path = arguments.get("file_path")
        
        if not file_path:
            raise ValueError("file_path is required")
        
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        file_extension = path.suffix.lower()
        
        # Route to appropriate reader based on file extension
        if file_extension == '.docx':
            return await self._read_word_document(arguments)
        elif file_extension in ['.xlsx', '.xls', '.xlsm']:
            return await self._read_excel_file(arguments)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}. Supported formats: {list(self.supported_formats.keys())}")
    
    async def _read_excel_file(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Read Excel file with comprehensive data extraction"""
        file_path = arguments.get("file_path")
        include_formatting = arguments.get("include_formatting", True)
        include_formulas = arguments.get("include_formulas", True)
        include_validation = arguments.get("include_validation", True)
        
        if not file_path:
            raise ValueError("file_path is required")
        
        try:
            result = self.excel_reader.read_excel_file(
                file_path, 
                include_formatting=include_formatting,
                include_formulas=include_formulas,
                include_validation=include_validation
            )
            return result
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "file_path": file_path
            }
    
    async def _read_word_document(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Read Word document"""
        file_path = arguments.get("file_path")
        include_formatting = arguments.get("include_formatting", False)
        
        if not file_path:
            raise ValueError("file_path is required")
        
        try:
            result = self.word_reader.read_word_document(file_path, include_formatting)
            return result
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "file_path": file_path
            }
    
    async def _list_supported_files(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """List all supported files in a directory"""
        directory_path = arguments.get("directory_path", ".")
        path = Path(directory_path)
        
        if not path.exists():
            raise FileNotFoundError(f"Directory not found: {directory_path}")
        
        if not path.is_dir():
            raise ValueError(f"Path is not a directory: {directory_path}")
        
        supported_files = []
        for file_path in path.rglob("*"):
            if file_path.is_file() and file_path.suffix.lower() in self.supported_formats:
                supported_files.append({
                    "name": file_path.name,
                    "path": str(file_path),
                    "format": file_path.suffix.lower(),
                    "format_description": self.supported_formats[file_path.suffix.lower()],
                    "size": file_path.stat().st_size,
                    "modified": file_path.stat().st_mtime
                })
        
        return {
            "success": True,
            "directory": str(path),
            "supported_files": supported_files,
            "count": len(supported_files),
            "supported_formats": list(self.supported_formats.keys())
        }
    
    async def _get_supported_formats(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Get list of all supported file formats"""
        return {
            "success": True,
            "supported_formats": self.supported_formats,
            "total_formats": len(self.supported_formats)
        }


async def main():
    """Main entry point for the MCP server"""
    server = AIFileBridgeServer()
    
    # Read from stdin and write to stdout
    while True:
        try:
            line = await asyncio.get_event_loop().run_in_executor(None, sys.stdin.readline)
            if not line:
                break
            
            request = json.loads(line.strip())
            response = await server.handle_request(request)
            print(json.dumps(response, ensure_ascii=False))
            sys.stdout.flush()
            
        except json.JSONDecodeError as e:
            error_response = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {
                    "code": -32700,
                    "message": f"Parse error: {str(e)}"
                }
            }
            print(json.dumps(error_response))
            sys.stdout.flush()
        except Exception as e:
            error_response = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {str(e)}"
                }
            }
            print(json.dumps(error_response))
            sys.stdout.flush()


if __name__ == "__main__":
    asyncio.run(main())
