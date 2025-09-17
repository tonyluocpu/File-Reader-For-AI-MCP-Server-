#!/usr/bin/env python3
"""
Example usage of the Word Reader MCP Server
This script demonstrates how to use the MCP server to read Word documents
"""

import json
import subprocess
import sys
from pathlib import Path

def example_read_document():
    """Example: Reading a Word document with formatting"""
    
    # Path to your Word document
    doc_path = "sample_document.docx"
    
    # Check if file exists
    if not Path(doc_path).exists():
        print(f"Sample document '{doc_path}' not found.")
        print("Please place a .docx file in this directory to test.")
        return
    
    print("=== Word Reader MCP Server Example ===")
    print(f"Reading document: {doc_path}")
    print()
    
    # Start the MCP server process
    try:
        process = subprocess.Popen(
            [sys.executable, "../word_reader_mcp_server.py"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=".."
        )
        
        # Initialize the server
        init_request = {
            "jsonrpc": "2.0",
            "id": 0,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "example-client",
                    "version": "1.0.0"
                }
            }
        }
        
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        
        # Read initialization response
        init_response = process.stdout.readline()
        print("✓ Server initialized successfully")
        
        # List available tools
        tools_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/list"
        }
        
        process.stdin.write(json.dumps(tools_request) + "\n")
        process.stdin.flush()
        
        tools_response = process.stdout.readline()
        print("✓ Tools listed successfully")
        
        # Read the document
        read_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/call",
            "params": {
                "name": "read_word_document",
                "arguments": {
                    "file_path": doc_path,
                    "include_formatting": True
                }
            }
        }
        
        process.stdin.write(json.dumps(read_request) + "\n")
        process.stdin.flush()
        
        # Read response
        response = process.stdout.readline()
        print("✓ Document read successfully")
        
        # Parse and display results
        try:
            response_data = json.loads(response)
            if "result" in response_data and "content" in response_data["result"]:
                content = json.loads(response_data["result"]["content"][0]["text"])
                
                if content.get("success"):
                    print(f"\n📄 Document: {content['file_path']}")
                    print(f"📊 Total paragraphs: {content['total_paragraphs']}")
                    print(f"📋 Total tables: {content['total_tables']}")
                    
                    print("\n📝 Sample paragraphs:")
                    print("-" * 40)
                    
                    # Show first 3 paragraphs as examples
                    for i, para in enumerate(content['paragraphs'][:3]):
                        print(f"\nParagraph {i+1}:")
                        print(f"Text: {para['text'][:100]}{'...' if len(para['text']) > 100 else ''}")
                        
                        if 'formatting' in para:
                            style = para['formatting'].get('style', 'Normal')
                            print(f"Style: {style}")
                    
                    if len(content['paragraphs']) > 3:
                        print(f"\n... and {len(content['paragraphs']) - 3} more paragraphs")
                        
                    # Show tables if any
                    if content['tables']:
                        print(f"\n📋 Tables found: {len(content['tables'])}")
                        for i, table in enumerate(content['tables'][:2]):  # Show first 2 tables
                            print(f"\nTable {i+1}:")
                            for row in table[:3]:  # Show first 3 rows
                                print(f"  {' | '.join(row)}")
                            if len(table) > 3:
                                print(f"  ... and {len(table) - 3} more rows")
                        
                else:
                    print(f"❌ Error reading document: {content.get('error', 'Unknown error')}")
            else:
                print("❌ Unexpected response format")
                
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing response: {e}")
        
        # Clean up
        process.stdin.close()
        process.terminate()
        process.wait()
        
        print("\n✅ Example completed successfully!")
        
    except Exception as e:
        print(f"❌ Error running example: {e}")

def example_list_documents():
    """Example: Listing Word documents in a directory"""
    
    print("\n=== Document Listing Example ===")
    
    try:
        process = subprocess.Popen(
            [sys.executable, "../word_reader_mcp_server.py"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=".."
        )
        
        # Initialize (reusing from previous example)
        init_request = {
            "jsonrpc": "2.0",
            "id": 0,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "example-client",
                    "version": "1.0.0"
                }
            }
        }
        
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        process.stdout.readline()  # Read init response
        
        # List documents in current directory
        list_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": "list_word_documents",
                "arguments": {
                    "directory_path": "."
                }
            }
        }
        
        process.stdin.write(json.dumps(list_request) + "\n")
        process.stdin.flush()
        
        response = process.stdout.readline()
        
        try:
            response_data = json.loads(response)
            if "result" in response_data and "content" in response_data["result"]:
                content = json.loads(response_data["result"]["content"][0]["text"])
                
                if content.get("success"):
                    print(f"📁 Directory: {content['directory']}")
                    print(f"📄 Found {content['count']} .docx files:")
                    
                    for doc in content['docx_files']:
                        size_kb = doc['size'] / 1024
                        print(f"  • {doc['name']} ({size_kb:.1f} KB)")
                else:
                    print(f"❌ Error listing documents: {content.get('error', 'Unknown error')}")
                    
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing response: {e}")
        
        process.stdin.close()
        process.terminate()
        process.wait()
        
    except Exception as e:
        print(f"❌ Error running list example: {e}")

if __name__ == "__main__":
    print("Word Reader MCP Server - Usage Examples")
    print("=" * 50)
    
    # Run examples
    example_read_document()
    example_list_documents()
    
    print("\n" + "=" * 50)
    print("For more examples, check the README.md file")
    print("To use with AI assistants, configure MCP settings as described in the README")
