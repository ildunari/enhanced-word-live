#!/bin/bash

# Enhanced Word MCP Server - Live Editing Startup Script
# This script automatically starts both the MCP server and Word Add-in development server

set -e

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ADDIN_DIR="$PROJECT_DIR/word-live-addin"

echo "🚀 Starting Enhanced Word MCP Server with Live Editing..."
echo "Project directory: $PROJECT_DIR"

# Function to check if a port is in use
check_port() {
    local port=$1
    if lsof -Pi :$port -sTCP:LISTEN -t >/dev/null ; then
        return 0  # Port is in use
    else
        return 1  # Port is free
    fi
}

# Function to start MCP server
start_mcp_server() {
    echo "📡 Starting MCP Server (WebSocket on port 8765)..."
    
    if check_port 8765; then
        echo "⚠️  Port 8765 is already in use - MCP server may already be running"
        echo "   If you need to restart, please stop the existing server first"
        return 1
    fi
    
    cd "$PROJECT_DIR"
    
    # Check if virtual environment exists and activate it
    if [ -d "venv" ]; then
        echo "   Activating virtual environment..."
        source venv/bin/activate
    elif [ -d ".venv" ]; then
        echo "   Activating virtual environment..."
        source .venv/bin/activate
    fi
    
    # Install dependencies if needed
    if [ -f "requirements.txt" ]; then
        echo "   Installing Python dependencies..."
        pip install -r requirements.txt >/dev/null 2>&1
    fi
    
    # Start MCP server in background
    echo "   Launching MCP server..."
    python -m word_document_server.main &
    MCP_PID=$!
    
    # Wait a moment for server to start
    sleep 3
    
    # Check if MCP server started successfully
    if check_port 8765; then
        echo "✅ MCP Server started successfully (PID: $MCP_PID)"
        echo "   WebSocket endpoint: ws://localhost:8765"
        return 0
    else
        echo "❌ Failed to start MCP server"
        return 1
    fi
}

# Function to start Word Add-in development server
start_addin_server() {
    echo "🔧 Starting Word Add-in Development Server (HTTPS on port 3000)..."
    
    if check_port 3000; then
        echo "⚠️  Port 3000 is already in use - Add-in server may already be running"
        echo "   If you need to restart, please stop the existing server first"
        return 1
    fi
    
    cd "$ADDIN_DIR"
    
    # Check if Node.js is installed
    if ! command -v node &> /dev/null; then
        echo "❌ Node.js is not installed. Please install Node.js and npm first."
        echo "   Visit: https://nodejs.org/"
        return 1
    fi
    
    # Install npm dependencies if node_modules doesn't exist
    if [ ! -d "node_modules" ]; then
        echo "   Installing npm dependencies..."
        npm install
    fi
    
    # Start Add-in development server in background
    echo "   Launching Add-in development server..."
    npm start &
    ADDIN_PID=$!
    
    # Wait for server to start
    sleep 5
    
    # Check if Add-in server started successfully
    if check_port 3000; then
        echo "✅ Add-in Development Server started successfully (PID: $ADDIN_PID)"
        echo "   HTTPS endpoint: https://localhost:3000"
        return 0
    else
        echo "❌ Failed to start Add-in development server"
        return 1
    fi
}

# Function to cleanup on exit
cleanup() {
    echo ""
    echo "🛑 Shutting down servers..."
    
    if [ ! -z "$MCP_PID" ]; then
        echo "   Stopping MCP server (PID: $MCP_PID)..."
        kill $MCP_PID 2>/dev/null || true
    fi
    
    if [ ! -z "$ADDIN_PID" ]; then
        echo "   Stopping Add-in server (PID: $ADDIN_PID)..."
        kill $ADDIN_PID 2>/dev/null || true
    fi
    
    # Kill any remaining processes on our ports
    lsof -ti:8765 | xargs kill -9 2>/dev/null || true
    lsof -ti:3000 | xargs kill -9 2>/dev/null || true
    
    echo "✅ Cleanup complete"
    exit 0
}

# Setup signal handlers for cleanup
trap cleanup EXIT INT TERM

# Main execution
echo "="*60
echo "🏗️  Enhanced Word MCP Server - Live Editing Setup"
echo "="*60

# Start MCP server
if start_mcp_server; then
    echo ""
else
    echo "❌ Failed to start MCP server. Exiting."
    exit 1
fi

# Start Add-in development server
if start_addin_server; then
    echo ""
else
    echo "❌ Failed to start Add-in server. Exiting."
    exit 1
fi

echo "="*60
echo "🎉 Live Editing System Ready!"
echo "="*60
echo ""
echo "📋 Next Steps:"
echo "1. Open Microsoft Word"
echo "2. Go to Insert → My Add-ins → Upload My Add-in"
echo "3. Select: $ADDIN_DIR/manifest.xml"
echo "4. Open a document and the Add-in will auto-connect!"
echo ""
echo "🔗 Endpoints:"
echo "   • MCP Server: ws://localhost:8765"
echo "   • Add-in Dev: https://localhost:3000"
echo ""
echo "📝 The Add-in will automatically detect and connect to the MCP server."
echo "   Look for the 'LIVE' status in the Word task pane."
echo ""
echo "⌨️  Press Ctrl+C to stop both servers"
echo ""

# Keep script running and wait for interrupt
while true; do
    sleep 1
done