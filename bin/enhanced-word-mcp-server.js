#!/usr/bin/env node

const { spawn, execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

// Get the directory where this package is installed
const packageDir = path.dirname(__dirname);

// Function to find the correct Python executable with mcp module
function findPythonWithMCP() {
  // Try environment variable first, then common paths
  const pythonPaths = [
    process.env.PYTHON_PATH,
    process.env.ENHANCED_WORD_PYTHON,
    'python3',
    'python',
    '/usr/bin/python3',
    '/usr/local/bin/python3'
  ].filter(Boolean); // Remove undefined values
  
  for (const pythonPath of pythonPaths) {
    try {
      // Check if this Python has the mcp module
      execSync(`${pythonPath} -c "import mcp; print('MCP found')"`, { 
        stdio: 'pipe', 
        timeout: 5000 
      });
      return pythonPath;
    } catch (err) {
      continue;
    }
  }
  
  throw new Error('No Python installation found with MCP module. Please install: pip install mcp');
}

try {
  const pythonExecutable = findPythonWithMCP();
  
  // Run the Python MCP server
  const python = spawn(pythonExecutable, ['-m', 'word_document_server.main'], {
    cwd: packageDir,
    stdio: 'inherit',
    env: { ...process.env }
  });

  python.on('close', (code) => {
    process.exit(code);
  });

  python.on('error', (err) => {
    console.error('Failed to start Enhanced Word MCP Server:', err.message);
    console.error('Make sure Python 3.11+ is installed with MCP module.');
    process.exit(1);
  });

} catch (err) {
  console.error('Setup Error:', err.message);
  console.error('Please install the MCP module: pip install mcp');
  console.error('Or ensure Python 3.11+ is in your PATH.');
  process.exit(1);
}
