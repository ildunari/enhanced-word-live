# Code Style and Conventions

## Python Code Style
- **Python Version**: 3.11+
- **Docstring Style**: Google-style docstrings
- **Function Naming**: snake_case
- **Class Naming**: PascalCase
- **Module Organization**: Tools organized by functionality in separate files

## File Organization
```
word_document_server/
├── main.py              # Main MCP server entry point
├── tools/               # MCP tool implementations
│   ├── document_tools.py    # Document creation/management
│   ├── content_tools.py     # Content manipulation
│   ├── format_tools.py      # Text formatting
│   ├── protection_tools.py  # Document protection
│   ├── footnote_tools.py    # Footnote/endnote tools
│   ├── review_tools.py      # Track changes/comments
│   └── section_tools.py     # Section management
├── core/                # Core functionality modules
├── utils/               # Utility functions
└── __init__.py
```

## Error Handling
- Custom exception classes defined in review_tools.py
- Comprehensive error handling for document operations
- Descriptive error messages for user guidance

## Documentation
- Each tool function has detailed docstrings
- Parameter descriptions include types and validation
- Return value documentation

## Testing
- Test file: `test_enhanced_features.py`
- Tests cover main functionality areas
- Manual testing with sample documents