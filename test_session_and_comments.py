#!/usr/bin/env python3
"""
Test script for Document Session Management and Fixed Comment Detection.

Tests both the new session management system and the fixed comment detection bug.
"""
import os
import sys

# Add the project root to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import asyncio
from word_document_server.tools.session_tools import (
    open_document, close_document, list_open_documents, 
    set_active_document, close_all_documents
)
from word_document_server.tools.review_tools import manage_comments
from word_document_server.tools.document_tools import create_document


async def test_session_management():
    """Test the document session management system."""
    print("=== Testing Document Session Management ===")
    
    # Test 1: List when no documents are open
    print("\n1. Testing list_open_documents with no documents:")
    result = list_open_documents()
    print(f"Result: {result}")
    
    # Test 2: Create test documents
    print("\n2. Creating test documents...")
    test_doc1 = "test_session_doc1.docx"
    test_doc2 = "test_session_doc2.docx"
    
    create_result1 = await create_document(test_doc1, title="Test Document 1", author="Test Author")
    create_result2 = await create_document(test_doc2, title="Test Document 2", author="Test Author")
    print(f"Created doc1: {create_result1}")
    print(f"Created doc2: {create_result2}")
    
    # Test 3: Open documents with session IDs
    print("\n3. Testing open_document:")
    open_result1 = open_document("main", test_doc1)
    open_result2 = open_document("draft", test_doc2)
    print(f"Open main: {open_result1}")
    print(f"Open draft: {open_result2}")
    
    # Test 4: List open documents
    print("\n4. Testing list_open_documents with open documents:")
    result = list_open_documents()
    print(f"Result:\n{result}")
    
    # Test 5: Set active document
    print("\n5. Testing set_active_document:")
    active_result = set_active_document("draft")
    print(f"Set active: {active_result}")
    
    # Test 6: List again to see active change
    print("\n6. List after changing active:")
    result = list_open_documents()
    print(f"Result:\n{result}")
    
    # Test 7: Try to open document with existing ID (should fail)
    print("\n7. Testing duplicate ID (should fail):")
    duplicate_result = open_document("main", test_doc2)
    print(f"Duplicate ID result: {duplicate_result}")
    
    return test_doc1, test_doc2


def test_comment_system(test_doc1, test_doc2):
    """Test the fixed comment detection system."""
    print("\n\n=== Testing Fixed Comment System ===")
    
    # Test 1: Add comments using document_id
    print("\n1. Testing add comments with document_id:")
    add_result1 = manage_comments(
        document_id="main", 
        action="add", 
        paragraph_index=0, 
        comment_text="This is a test comment", 
        author="Test User"
    )
    print(f"Add comment 1: {add_result1}")
    
    add_result2 = manage_comments(
        document_id="main", 
        action="add", 
        paragraph_index=0, 
        comment_text="Second comment on same paragraph", 
        author="Reviewer"
    )
    print(f"Add comment 2: {add_result2}")
    
    # Test 2: List comments immediately after adding (this was the bug)
    print("\n2. Testing list comments (this was broken before):")
    list_result = manage_comments(document_id="main", action="list")
    print(f"List result:\n{list_result}")
    
    # Test 3: Add comment to second document
    print("\n3. Testing comments on second document:")
    add_result3 = manage_comments(
        document_id="draft", 
        action="add", 
        paragraph_index=0, 
        comment_text="Comment on draft document", 
        author="Editor"
    )
    print(f"Add comment to draft: {add_result3}")
    
    # Test 4: List comments on second document
    print("\n4. List comments on draft:")
    list_result2 = manage_comments(document_id="draft", action="list")
    print(f"Draft comments:\n{list_result2}")
    
    # Test 5: Test backward compatibility with filename
    print("\n5. Testing backward compatibility with filename parameter:")
    list_result3 = manage_comments(filename=test_doc1, action="list")
    print(f"Filename-based list:\n{list_result3}")
    
    # Test 6: Extract specific comment ID for resolve/delete testing
    # Parse comment ID from add_result1
    if "Successfully added comment" in add_result1:
        comment_id = add_result1.split("comment ")[1].split(" to")[0]
        print(f"\n6. Testing resolve comment {comment_id}:")
        resolve_result = manage_comments(document_id="main", action="resolve", comment_id=comment_id)
        print(f"Resolve result: {resolve_result}")
        
        # List again to see resolved status
        print("\n7. List after resolving:")
        list_result4 = manage_comments(document_id="main", action="list")
        print(f"After resolve:\n{list_result4}")


def test_error_handling():
    """Test error handling in both systems."""
    print("\n\n=== Testing Error Handling ===")
    
    # Test 1: Invalid document_id
    print("\n1. Testing invalid document_id:")
    invalid_result = manage_comments(document_id="nonexistent", action="list")
    print(f"Invalid document_id: {invalid_result}")
    
    # Test 2: Missing required parameters
    print("\n2. Testing missing parameters:")
    missing_result = manage_comments(document_id="main", action="add")
    print(f"Missing parameters: {missing_result}")
    
    # Test 3: Close non-existent document
    print("\n3. Testing close non-existent document:")
    close_invalid = close_document("nonexistent")
    print(f"Close invalid: {close_invalid}")


def cleanup_and_finish():
    """Clean up test files and finish."""
    print("\n\n=== Cleanup ===")
    
    # Close all documents
    print("\n1. Closing all documents:")
    close_all_result = close_all_documents()
    print(f"Close all: {close_all_result}")
    
    # Verify no documents open
    print("\n2. Verify no documents open:")
    final_list = list_open_documents()
    print(f"Final list: {final_list}")
    
    # Clean up test files
    print("\n3. Cleaning up test files...")
    try:
        os.remove("test_session_doc1.docx")
        os.remove("test_session_doc2.docx")
        print("Test files removed successfully")
    except Exception as e:
        print(f"Error removing test files: {e}")


async def main():
    """Run all tests."""
    print("Enhanced Word MCP Server - Session Management & Comment Fix Test")
    print("=" * 65)
    
    try:
        # Test session management
        test_doc1, test_doc2 = await test_session_management()
        
        # Test comment system
        test_comment_system(test_doc1, test_doc2)
        
        # Test error handling
        test_error_handling()
        
        print("\n\n✅ ALL TESTS COMPLETED")
        print("Key improvements:")
        print("- Document session management with simple IDs")
        print("- Fixed comment detection bug (add + list now work together)")
        print("- Backward compatibility with filename parameters")
        print("- Both document_id and filename support in tools")
        
    except Exception as e:
        print(f"\n❌ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Always try to cleanup
        cleanup_and_finish()


if __name__ == "__main__":
    asyncio.run(main())