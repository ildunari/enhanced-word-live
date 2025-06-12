#!/usr/bin/env python3
"""
Test script to verify the fixed regex batch processing.
This script will run the fixed enhanced_search_and_replace function on our test document
and verify that the corruption is now fixed.
"""

import sys
import os
from docx import Document
from pathlib import Path
import shutil

# Add the project directory to Python path
project_dir = Path(__file__).parent
sys.path.insert(0, str(project_dir))

# Import the fixed function
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

async def test_fixed_regex():
    """Test the fixed regex processing."""
    
    test_file = "regex_test_document.docx"
    if not os.path.exists(test_file):
        print(f"❌ Test file {test_file} not found. Run create_test_document.py first.")
        return
    
    # Make a fresh copy for testing the fix
    test_copy = "regex_test_document_fixed.docx"
    shutil.copy2(test_file, test_copy)
    
    print("🔧 Testing FIXED regex batch processing...")
    print(f"📄 Using test file: {test_copy}")
    
    # Read initial state
    print("\n📋 Reading initial document state...")
    initial_state = read_document_paragraphs(test_copy)
    print_document_state(initial_state, "BEFORE: Original Document State")
    
    # Run the FIXED regex replacement
    print("\n🔄 Running FIXED enhanced_search_and_replace...")
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
        import traceback
        traceback.print_exc()
        return
    
    # Read final state
    print("\n📋 Reading final document state...")
    final_state = read_document_paragraphs(test_copy)
    print_document_state(final_state, "AFTER: Document State After FIXED Regex Processing")
    
    # Compare states and verify fixes
    print("\n🔍 FIX VERIFICATION:")
    print("="*60)
    
    success_count = 0
    total_paragraphs = len(initial_state)
    
    for i, (initial, final) in enumerate(zip(initial_state, final_state)):
        if initial['text'] == final['text']:
            print(f"\n✅ PARAGRAPH {i}: Text structure PRESERVED")
            print(f"   Text: '{final['text']}'")
            
            # Check if formatting was applied correctly to curly brace content
            bold_runs = [run for run in final['runs'] if run['bold']]
            curly_content = [run for run in final['runs'] if '{' in run['text'] and '}' in run['text']]
            bold_curly = [run for run in curly_content if run['bold']]
            
            if curly_content:
                if len(bold_curly) == len(curly_content):
                    print(f"   ✅ Bold formatting applied correctly to {len(bold_curly)} curly brace items")
                else:
                    print(f"   ⚠️  Partial bold formatting: {len(bold_curly)}/{len(curly_content)} items")
            
            success_count += 1
        else:
            print(f"\n❌ PARAGRAPH {i}: STILL CORRUPTED")
            print(f"   Before: '{initial['text']}'")
            print(f"   After:  '{final['text']}'")
    
    # Summary
    print(f"\n{'='*60}")
    print(f"🎯 FIX SUCCESS RATE: {success_count}/{total_paragraphs} paragraphs preserved")
    print(f"   ({(success_count/total_paragraphs)*100:.1f}% success rate)")
    
    if success_count == total_paragraphs:
        print("🎉 ALL CORRUPTION ISSUES FIXED! ✅")
    else:
        print(f"⚠️  {total_paragraphs - success_count} paragraphs still have issues")
    
    return success_count == total_paragraphs

def main():
    """Main function to run the test."""
    import asyncio
    success = asyncio.run(test_fixed_regex())
    return success

if __name__ == "__main__":
    main()