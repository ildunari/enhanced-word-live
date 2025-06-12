# Critical Bugs Found in Previous Version - Debug Checklist

## 🚨 CRITICAL BUGS TO INVESTIGATE IN NEW VERSION

### 1. FORMATTING OVER-APPLICATION BUG (Critical)
- **Issue**: enhanced_search_and_replace applies formatting to entire paragraphs instead of matched text only
- **Root Cause**: Character range calculation errors in formatting logic
- **Check**: Look for paragraph-level vs character-level formatting application
- **Test Case**: Search for "PCL" and apply red color - should only affect "PCL" instances, not entire paragraphs

### 2. COMMENT DETECTION BUG (High Priority)  
- **Issue**: extract_comments always returns "No comments found" even when comments exist
- **Root Cause**: Comment parsing logic or API usage issue
- **Check**: Compare add_comment (works) vs extract_comments (broken) implementations
- **Test Case**: Add comment via add_comment, then try extract_comments - should find the added comment

### 3. ASYNC/AWAIT ERRORS (Medium Priority)
- **Issue**: generate_review_summary, get_author_specific_changes fail with "object str can't be used in 'await' expression"
- **Root Cause**: Incorrect await usage on string objects instead of coroutines
- **Check**: Review async function definitions and await patterns
- **Test Case**: Call these functions and ensure no async errors

### 4. INCONSISTENT FORMATTING APPLICATION
- **Issue**: format_specific_words applies formatting to some but not all instances
- **Root Cause**: Document state/indexing issues after modifications
- **Check**: Text indexing refresh logic after document changes
- **Test Case**: Format word that appears multiple times - all instances should be formatted

## ✅ TOOLS THAT WORKED PERFECTLY (Reference)
- search_and_replace (text-only)
- format_text (character position approach)
- All content creation tools
- All document analysis tools