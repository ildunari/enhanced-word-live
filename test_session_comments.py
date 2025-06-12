#!/usr/bin/env python3
"""
Quick test for session management + comments integration.
"""
import os
import sys
import asyncio

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_document_server.tools.session_tools import session_manager
from word_document_server.tools.review_tools import manage_comments
from word_document_server.tools.document_tools import create_document
from word_document_server.tools.content_tools import add_text_content


async def test_session_comments_integration():
    """Test session management with comment functionality."""
    print("=== Testing Session + Comments Integration ===")
    
    # Create a test document with content
    test_doc = "test_session_comments.docx"
    
    print("1. Creating test document with content...")
    create_result = await create_document(test_doc, title="Session Comment Test", author="Test User")
    print(f"Created: {create_result}")
    
    # Add some content to comment on
    content_result = await add_text_content(test_doc, "This is the first paragraph for testing comments.")
    print(f"Added content: {content_result}")
    
    content_result2 = await add_text_content(test_doc, "This is the second paragraph with important information.")
    print(f"Added content: {content_result2}")
    
    # Open document in session
    print("\n2. Opening document in session...")
    open_result = session_manager("open", document_id="test_doc", file_path=test_doc)
    print(f"Opened: {open_result}")
    
    # List sessions
    print("\n3. Listing open sessions...")
    list_result = session_manager("list")
    print(f"Sessions:\n{list_result}")
    
    # Add comments using document_id
    print("\n4. Adding comments using document_id...")
    comment1 = manage_comments(
        document_id="test_doc", 
        action="add", 
        paragraph_index=0, 
        comment_text="This needs clarification",
        author="Reviewer 1"
    )
    print(f"Comment 1: {comment1}")
    
    comment2 = manage_comments(
        document_id="test_doc", 
        action="add", 
        paragraph_index=1, 
        comment_text="Important point here",
        author="Reviewer 2"
    )
    print(f"Comment 2: {comment2}")
    
    # List comments using document_id
    print("\n5. Listing comments using document_id...")
    list_comments = manage_comments(document_id="test_doc", action="list")
    print(f"Comments:\n{list_comments}")
    
    # Test backward compatibility with filename
    print("\n6. Testing backward compatibility with filename...")
    list_comments_filename = manage_comments(filename=test_doc, action="list")
    print(f"Comments via filename:\n{list_comments_filename}")
    
    # Close session
    print("\n7. Closing session...")
    close_result = session_manager("close", document_id="test_doc")
    print(f"Closed: {close_result}")
    
    # Cleanup
    print("\n8. Cleanup...")
    try:
        os.remove(test_doc)
        print(f"Removed {test_doc}")
    except Exception as e:
        print(f"Error removing {test_doc}: {e}")
    
    print("\n✅ Session + Comments integration test completed!")


if __name__ == "__main__":
    asyncio.run(test_session_comments_integration())