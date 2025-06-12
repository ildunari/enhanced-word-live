#!/usr/bin/env python3
"""
Test script to demonstrate the regex batch processing corruption.
This script will run the enhanced_search_and_replace function on our test document
and capture the before/after states to show exactly what corruption occurs.
"""

import sys
import os
from docx import Document
from pathlib import Path

# Add the project directory to Python path
project_dir = Path(__file__).parent
sys.path.insert(0, str(project_dir))

# Import the problematic function
from word_document_server.tools.content_tools import enhanced_search_and_replace

def read_document_paragraphs(filename):
    """Read all paragraphs from a document and return as list."""
    doc = Document(filename)
    paragraphs = []
    
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():  # Skip empty paragraphs
            paragraphs.append({
                'index': i,
                'text': para.text,
                'runs': [{'text': run.text, 'bold': run.bold, 'italic': run.italic} 
                        for run in para.runs if run.text]
            })
    
    return paragraphs

def print_document_state(paragraphs, title):
    """Print the document state in a readable format."""
    print(f"\n{'='*60}")
    print(f"{title}")
    print('='*60)
    
    for para in paragraphs:
        print(f"\nParagraph {para['index']}:")
        print(f"  Text: '{para['text']}'")
        print(f"  Runs: {len(para['runs'])}")
        for j, run in enumerate(para['runs']):
            formatting = []
            if run['bold']: formatting.append('BOLD')
            if run['italic']: formatting.append('ITALIC')
            fmt_str = f" [{', '.join(formatting)}]" if formatting else ""
            print(f"    Run {j}: '{run['text']}'{fmt_str}")

async def test_regex_corruption():
    """Test the regex corruption by running enhanced_search_and_replace."""
    
    test_file = "regex_test_document.docx"
    if not os.path.exists(test_file):
        print(f"❌ Test file {test_file} not found. Run create_test_document.py first.")
        return
    
    # Make a copy for testing
    import shutil
    test_copy = "regex_test_document_copy.docx"
    shutil.copy2(test_file, test_copy)
    
    print("🔍 Testing regex batch processing corruption...")
    print(f"📄 Using test file: {test_copy}")
    
    # Read initial state
    print("\n📋 Reading initial document state...")
    initial_state = read_document_paragraphs(test_copy)
    print_document_state(initial_state, "BEFORE: Original Document State")
    
    # Run the problematic regex replacement
    print("\n🔄 Running enhanced_search_and_replace...")
    print("   Pattern: \\{[^}]+\\}")
    print("   Replace: (same text)")
    print("   Formatting: Bold = True")
    
    try:
        result = await enhanced_search_and_replace(
            filename=test_copy,
            find_text=r"\{[^}]+\}",  # Match any text in curly braces
            replace_text=r"\g<0>",   # Replace with the same text (no text change)
            use_regex=True,
            apply_formatting=True,
            bold=True,
            match_case=True
        )
        
        print(f"✅ Function result: {result}")
        
    except Exception as e:
        print(f"❌ Error during replacement: {e}")
        return
    
    # Read final state
    print("\n📋 Reading final document state...")
    final_state = read_document_paragraphs(test_copy)
    print_document_state(final_state, "AFTER: Document State After Regex Processing")
    
    # Compare states
    print("\n🔍 CORRUPTION ANALYSIS:")
    print("="*60)
    
    for i, (initial, final) in enumerate(zip(initial_state, final_state)):
        if initial['text'] != final['text']:
            print(f"\n❌ PARAGRAPH {i} CORRUPTED:")
            print(f"   Before: '{initial['text']}'")
            print(f"   After:  '{final['text']}'")
            print(f"   Runs changed: {len(initial['runs'])} → {len(final['runs'])}")
        else:
            print(f"\n✅ Paragraph {i}: Text unchanged")
            # Check if formatting was applied correctly
            bold_runs = [run for run in final['runs'] if run['bold']]
            if bold_runs:
                print(f"   Bold formatting applied to {len(bold_runs)} runs")

def main():
    """Main function to run the test."""
    import asyncio
    asyncio.run(test_regex_corruption())

if __name__ == "__main__":
    main()