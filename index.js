/**
 * Enhanced Word MCP Server - Node.js Entry Point
 * 
 * This is a Node.js wrapper for the Python-based Enhanced Word MCP Server.
 * The actual server logic is implemented in Python using the word_document_server module.
 */

const { spawn } = require('child_process');
const path = require('path');

/**
 * Start the Enhanced Word MCP Server
 * @param {Object} options - Configuration options
 * @returns {ChildProcess} The spawned Python process
 */
function startServer(options = {}) {
  const packageDir = __dirname;
  
  const python = spawn('python', ['-m', 'word_document_server.main'], {
    cwd: packageDir,
    stdio: options.stdio || 'inherit',
    env: { ...process.env, ...options.env }
  });

  return python;
}

module.exports = {
  startServer
};

// If this file is run directly, start the server
if (require.main === module) {
  const server = startServer();
  
  server.on('close', (code) => {
    process.exit(code);
  });
  
  server.on('error', (err) => {
    console.error('Failed to start Enhanced Word MCP Server:', err.message);
    console.error('Make sure Python 3.11+ is installed and requirements are met.');
    process.exit(1);
  });
}
