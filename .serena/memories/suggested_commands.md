# Suggested Commands for Enhanced Word MCP Server

## Development Commands
```bash
# Install Python dependencies
pip install -r requirements.txt

# Install Python package in development mode
pip install -e .

# Test the server functionality
python test_enhanced_features.py

# Run the server directly (Python)
python -m word_document_server.main

# Run via NPM scripts
npm run start
npm run test
npm run install-deps
```

## MCP Server Installation
```bash
# Install via NPM (for end users)
npx enhanced-word-mcp-server

# Install locally for development
npm install -g .
```

## Python Requirements Check
```bash
# Check if MCP module is available
python -c "import mcp; print('MCP found')"

# Install MCP if missing
pip install mcp
```

## System Commands (macOS/Darwin)
- `ls` - List directory contents
- `cd` - Change directory  
- `grep` - Search text patterns
- `find` - Find files
- `python3` - Python interpreter
- `pip` - Python package manager
- `npm` - Node.js package manager

## Git Commands
```bash
git status
git add .
git commit -m "message"
git push
```