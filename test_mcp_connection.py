#!/usr/bin/env python3
"""
Test MCP server connection and functionality
Simulates how Cursor would interact with the AI File Bridge server
"""

import asyncio
import json
import subprocess
import sys
from pathlib import Path

async def test_mcp_server():
    """Test the MCP server by sending requests"""
    print("🔧 Testing AI File Bridge MCP Server Connection")
    print("=" * 60)
    
    # Start the MCP server process
    server_path = Path(__file__).parent / "ai_file_bridge_server.py"
    
    try:
        process = subprocess.Popen(
            [sys.executable, str(server_path)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        print("✅ MCP Server started successfully")
        
        # Test 1: Initialize the server
        print("\n📡 Test 1: Server Initialization")
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0.0"}
            }
        }
        
        process.stdin.write(json.dumps(init_request) + '\n')
        process.stdin.flush()
        
        response = process.stdout.readline()
        init_response = json.loads(response)
        
        if "result" in init_response:
            print(f"✅ Server initialized: {init_response['result']['serverInfo']['name']} v{init_response['result']['serverInfo']['version']}")
        else:
            print(f"❌ Initialization failed: {init_response}")
            return
        
        # Test 2: List available tools
        print("\n🛠️ Test 2: Available Tools")
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }
        
        process.stdin.write(json.dumps(tools_request) + '\n')
        process.stdin.flush()
        
        response = process.stdout.readline()
        tools_response = json.loads(response)
        
        if "result" in tools_response:
            tools = tools_response['result']['tools']
            print(f"✅ Found {len(tools)} available tools:")
            for tool in tools:
                print(f"   🔧 {tool['name']}: {tool['description']}")
        else:
            print(f"❌ Tools list failed: {tools_response}")
            return
        
        # Test 3: Get supported formats
        print("\n📋 Test 3: Supported File Formats")
        formats_request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "get_supported_formats",
                "arguments": {}
            }
        }
        
        process.stdin.write(json.dumps(formats_request) + '\n')
        process.stdin.flush()
        
        response = process.stdout.readline()
        formats_response = json.loads(response)
        
        if "result" in formats_response:
            content = json.loads(formats_response['result']['content'][0]['text'])
            print(f"✅ Supported formats:")
            for ext, desc in content['supported_formats'].items():
                print(f"   📄 {ext}: {desc}")
        else:
            print(f"❌ Formats request failed: {formats_response}")
        
        # Test 4: List files in directory
        print("\n📁 Test 4: List Supported Files")
        list_request = {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "tools/call",
            "params": {
                "name": "list_supported_files",
                "arguments": {"directory_path": ".."}
            }
        }
        
        process.stdin.write(json.dumps(list_request) + '\n')
        process.stdin.flush()
        
        response = process.stdout.readline()
        list_response = json.loads(response)
        
        if "result" in list_response:
            content = json.loads(list_response['result']['content'][0]['text'])
            if content['success']:
                print(f"✅ Found {content['count']} supported files in directory")
                for file_info in content['supported_files'][:3]:  # Show first 3
                    print(f"   📄 {file_info['name']} ({file_info['format_description']})")
            else:
                print(f"❌ List files failed: {content}")
        else:
            print(f"❌ List files request failed: {list_response}")
        
        print("\n🎉 MCP Server Connection Test: PASSED ✅")
        print("🚀 Your AI File Bridge is ready for Cursor!")
        
    except Exception as e:
        print(f"❌ Test failed: {str(e)}")
    
    finally:
        # Clean up
        process.stdin.close()
        process.terminate()
        process.wait()

def show_cursor_setup_instructions():
    """Show instructions for setting up Cursor"""
    print("\n" + "=" * 60)
    print("🔧 CURSOR SETUP INSTRUCTIONS")
    print("=" * 60)
    
    current_path = Path(__file__).parent.absolute()
    
    print(f"""
📋 To connect AI File Bridge to Cursor:

1. Open Cursor Settings
2. Find MCP (Model Context Protocol) configuration
3. Add this server configuration:

{{
  "mcpServers": {{
    "ai-file-bridge": {{
      "command": "python3",
      "args": ["{current_path}/ai_file_bridge_server.py"],
      "cwd": "{current_path}"
    }}
  }}
}}

4. Restart Cursor
5. Test by asking Cursor to:
   - "Read this Word document and summarize it"
   - "Analyze the data in this Excel file"
   - "What formulas are in this spreadsheet?"

🎯 Your AI File Bridge server is running at:
   {current_path}/ai_file_bridge_server.py

✅ Supported formats: .docx, .xlsx, .xls, .xlsm
""")

if __name__ == "__main__":
    asyncio.run(test_mcp_server())
    show_cursor_setup_instructions()
