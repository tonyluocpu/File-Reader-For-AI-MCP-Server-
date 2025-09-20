#!/usr/bin/env python3
"""
Test script for the Word Reader MCP Server
"""

import json
import subprocess
import sys
from pathlib import Path

def test_mcp_server():
    """Test the MCP server with the Word document"""
    
    # Path to the Word document (replace with your actual path)
    doc_path = "/path/to/your/document.docx"
    
    # Test request to read the Word document
    test_request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "read_word_document",
            "arguments": {
                "file_path": doc_path,
                "include_formatting": True
            }
        }
    }
    
    try:
        # Start the MCP server process
        process = subprocess.Popen(
            [sys.executable, "word_reader_mcp_server.py"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Send initialization request first
        init_request = {
            "jsonrpc": "2.0",
            "id": 0,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }
        
        # Send initialization
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        
        # Read initialization response
        init_response = process.stdout.readline()
        print("Initialization response:", init_response.strip())
        
        # Send tools/list request
        tools_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/list"
        }
        
        process.stdin.write(json.dumps(tools_request) + "\n")
        process.stdin.flush()
        
        tools_response = process.stdout.readline()
        print("Tools list response:", tools_response.strip())
        
        # Send test request
        process.stdin.write(json.dumps(test_request) + "\n")
        process.stdin.flush()
        
        # Read response
        response = process.stdout.readline()
        print("Document reading response:", response.strip())
        
        # Parse and display the content
        try:
            response_data = json.loads(response)
            if "result" in response_data and "content" in response_data["result"]:
                content = json.loads(response_data["result"]["content"][0]["text"])
                print("\n" + "="*50)
                print("DOCUMENT CONTENT EXTRACTED:")
                print("="*50)
                
                if content.get("success"):
                    print(f"File: {content['file_path']}")
                    print(f"Total paragraphs: {content['total_paragraphs']}")
                    print(f"Total tables: {content['total_tables']}")
                    print("\nPARAGRAPHS:")
                    print("-" * 30)
                    
                    for i, para in enumerate(content['paragraphs'][:10]):  # Show first 10 paragraphs
                        print(f"\nParagraph {i+1}:")
                        print(f"Text: {para['text']}")
                        if 'formatting' in para:
                            print(f"Style: {para['formatting'].get('style', 'N/A')}")
                        
                        if 'runs' in para:
                            for run in para['runs']:
                                if run.get('bold') or run.get('italic'):
                                    print(f"  - Run: '{run['text']}' (Bold: {run.get('bold')}, Italic: {run.get('italic')})")
                    
                    if len(content['paragraphs']) > 10:
                        print(f"\n... and {len(content['paragraphs']) - 10} more paragraphs")
                        
                else:
                    print(f"Error: {content.get('error', 'Unknown error')}")
                    
        except json.JSONDecodeError as e:
            print(f"Error parsing response: {e}")
        
        # Clean up
        process.stdin.close()
        process.terminate()
        process.wait()
        
    except Exception as e:
        print(f"Error testing MCP server: {e}")

if __name__ == "__main__":
    test_mcp_server()
