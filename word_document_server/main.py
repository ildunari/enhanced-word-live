"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
"""

import os
import sys
import asyncio
import json
import threading
import websockets
from mcp.server.fastmcp import FastMCP
from word_document_server.tools import (
    document_tools,
    content_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    review_tools,
    section_tools,
    session_tools
)
from word_document_server.session_manager import get_session_manager



# Initialize FastMCP server
mcp = FastMCP("word-document-server")

def register_tools():
    """Register all tools with the MCP server - CONSOLIDATED VERSION WITH SESSION MANAGEMENT."""
    
    # ========== SESSION MANAGEMENT TOOLS (CONSOLIDATED) ==========
    # Unified session management (replaces 5 individual tools)
    mcp.tool()(session_tools.session_manager)
    
    # ========== CONSOLIDATED TOOLS (NEW) ==========
    # These replace multiple existing tools with enhanced functionality
    
    # Unified text extraction (replaces get_document_text, get_paragraph_text_from_document, find_text_in_document)
    mcp.tool()(document_tools.get_text)
    
    # Unified track changes management (replaces accept_all_changes, reject_all_changes)
    mcp.tool()(review_tools.manage_track_changes)
    
    # Unified note addition (replaces add_footnote_to_document, add_endnote_to_document)
    mcp.tool()(footnote_tools.add_note)
    
    # Unified text content addition (replaces add_paragraph, add_heading)
    mcp.tool()(content_tools.add_text_content)
    
    # Unified section extraction (replaces extract_sections_by_heading, extract_section_content)
    mcp.tool()(section_tools.get_sections)
    
    # Unified protection management (replaces protect_document, unprotect_document)
    mcp.tool()(protection_tools.manage_protection)
    
    # Enhanced comment management (replaces extract_comments with full lifecycle management)
    mcp.tool()(review_tools.manage_comments)
    
    # ========== CONSOLIDATED DOCUMENT TOOLS (NEW) ==========
    # Unified document utilities (replaces 3 individual tools)
    mcp.tool()(document_tools.document_utility)
    
    # ========== ESSENTIAL DOCUMENT TOOLS (7) ==========
    # Core document management that cannot be consolidated
    mcp.tool()(document_tools.create_document)
    mcp.tool()(document_tools.copy_document)
    mcp.tool()(document_tools.merge_documents)
    mcp.tool()(content_tools.enhanced_search_and_replace)
    mcp.tool()(content_tools.add_table)
    mcp.tool()(content_tools.add_picture)
    mcp.tool()(extended_document_tools.convert_to_pdf)
    
    # ========== CONSOLIDATED FORMATTING TOOLS (NEW) ==========
    # Unified document formatting (replaces 2 individual tools)
    mcp.tool()(content_tools.format_document)
    
    # ========== ADVANCED FEATURES (5) ==========
    # Specialized functionality for advanced use cases
    mcp.tool()(review_tools.extract_track_changes)
    mcp.tool()(review_tools.generate_review_summary)
    mcp.tool()(section_tools.generate_table_of_contents)
    mcp.tool()(protection_tools.add_digital_signature)
    mcp.tool()(protection_tools.verify_document)

    
    # ========== LEGACY COMPATIBILITY (OPTIONAL) ==========
    # These maintain backwards compatibility - can be removed after transition
    # Uncomment if you need backwards compatibility during transition period
    
    # Legacy text extraction tools (now replaced by get_text)
    # mcp.tool()(document_tools.get_document_text)
    # mcp.tool()(extended_document_tools.get_paragraph_text_from_document)
    # mcp.tool()(extended_document_tools.find_text_in_document)
    
    # Legacy track changes tools (now replaced by manage_track_changes)
    # mcp.tool()(review_tools.accept_all_changes)
    # mcp.tool()(review_tools.reject_all_changes)
    
    # Legacy note tools (now replaced by add_note)
    # mcp.tool()(footnote_tools.add_footnote_to_document)
    # mcp.tool()(footnote_tools.add_endnote_to_document)
    
    # Legacy content tools (now replaced by add_text_content)
    # mcp.tool()(content_tools.add_paragraph)
    # mcp.tool()(content_tools.add_heading)
    
    # Legacy section tools (now replaced by get_sections)
    # mcp.tool()(section_tools.extract_sections_by_heading)
    # mcp.tool()(section_tools.extract_section_content)
    
    # Legacy protection tools (now replaced by manage_protection)
    # mcp.tool()(protection_tools.protect_document)
    # mcp.tool()(protection_tools.unprotect_document)
    
    # Legacy basic search (now replaced by enhanced_search_and_replace and get_text with search scope)
    # mcp.tool()(content_tools.search_and_replace)






async def websocket_handler(websocket, path):
    """
    Handle incoming WebSocket connections from Word Add-ins.
    
    This function manages the communication between the Word Add-in and the 
    document session manager for live editing capabilities.
    """
    print(f"[WebSocket] Client connected from {websocket.remote_address}")
    session_manager = get_session_manager()
    registered_document_id = None
    
    try:
        async for message in websocket:
            try:
                data = json.loads(message)
                msg_type = data.get("type")
                
                if msg_type == "register":
                    # Word Add-in is registering for live editing
                    file_path = data.get("path")
                    if file_path:
                        # Find the document by file path
                        document_id = None
                        for doc_id, handle in session_manager._documents.items():
                            if handle.file_path == file_path:
                                document_id = doc_id
                                break
                        
                        if document_id:
                            result = session_manager.register_live_connection(document_id, websocket)
                            registered_document_id = document_id
                            print(f"[WebSocket] {result}")
                        else:
                            print(f"[WebSocket] No session found for document path: {file_path}")
                            # TODO: Could auto-create session here if needed
                            
                elif msg_type == "response":
                    # Response from Word Add-in to a previous request
                    correlation_id = data.get("correlation_id")
                    if correlation_id:
                        session_manager.handle_live_response(websocket, correlation_id, data)
                    else:
                        print(f"[WebSocket] Received response without correlation_id")
                        
                else:
                    print(f"[WebSocket] Received unknown message type: {msg_type}")
                    
            except json.JSONDecodeError as e:
                print(f"[WebSocket] Failed to parse JSON message: {e}")
            except Exception as e:
                print(f"[WebSocket] Error handling message: {e}")

    except websockets.exceptions.ConnectionClosed:
        # Normal disconnect
        pass
    except Exception as e:
        print(f"[WebSocket] Connection error: {e}")
    finally:
        # Clean up the live connection when WebSocket disconnects
        if registered_document_id:
            session_manager.unregister_live_connection(registered_document_id)
        print(f"[WebSocket] Client from {websocket.remote_address} disconnected")


def run_websocket_server():
    """
    Start the WebSocket server for live document editing.
    
    This runs in a separate thread so it doesn't block the main MCP transport.
    The WebSocket server listens on localhost:8765 for connections from Word Add-ins.
    """
    try:
        # Create a new event loop for this thread
        asyncio.set_event_loop(asyncio.new_event_loop())
        loop = asyncio.get_event_loop()
        
        # Start the WebSocket server
        start_server = websockets.serve(websocket_handler, "localhost", 8765)
        
        print("[WebSocket] Starting WebSocket server on ws://localhost:8765")
        loop.run_until_complete(start_server)
        loop.run_forever()
        
    except Exception as e:
        print(f"[WebSocket] Failed to start WebSocket server: {e}")

def run_server():
    """Run the Word Document MCP Server with live editing support."""
    # Register all tools
    register_tools()
    
    # Start WebSocket server in a background thread for live editing
    print("[Main] Starting Word Document MCP Server with live editing support...")
    websocket_thread = threading.Thread(target=run_websocket_server, daemon=True)
    websocket_thread.start()
    print("[Main] WebSocket server started for live document editing")
    
    # Run the main MCP server
    print("[Main] Starting MCP server...")
    mcp.run(transport='stdio')
    return mcp


if __name__ == "__main__":
    run_server()
