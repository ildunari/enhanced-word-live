# 🚀 Quick Start Guide - Enhanced Word MCP Server with Live Editing

## One-Click Setup (Automatic) ✨

### 1. Run the Startup Script

**macOS/Linux:**
```bash
./start-live-editing.sh
```

**Windows:**
```batch
start-live-editing.bat
```

This automatically:
- ✅ Starts the MCP server (WebSocket on port 8765)
- ✅ Starts the Word Add-in development server (HTTPS on port 3000)  
- ✅ Installs any missing dependencies
- ✅ Provides helpful setup instructions

### 2. Install the Add-in in Word

1. **Open Microsoft Word**
2. **Go to Insert → My Add-ins → Upload My Add-in**
3. **Select:** `word-live-addin/manifest.xml`
4. **Open any document** - the Add-in appears in the task pane

### 3. Automatic Connection 🤖

The Add-in automatically:
- 🔍 **Detects** the MCP server every 2 seconds
- ⚡ **Connects** as soon as the server is available
- 🔄 **Reconnects** if the connection is lost
- 💚 **Shows "LIVE"** status when connected

**No manual clicking required!**

---

## How It Works

### 🏗️ Architecture
```
Word Document ↔ Office Add-in ↔ WebSocket ↔ MCP Server ↔ AI Agent
```

### 🔄 Connection Flow
1. **Add-in loads** in Word task pane
2. **Auto-detection starts** - searches for MCP server
3. **Connection established** when server is found
4. **Status shows "LIVE"** - ready for AI editing
5. **Auto-reconnect** if connection drops

### ⚡ Live Editing Features
- **Real-time text replacement** with formatting
- **Live content extraction** for analysis
- **Optimized Office.js commands** for speed
- **Fallback to file mode** when offline

---

## Using Live Editing

### Prerequisites
- **MCP server running** (via startup script)
- **Add-in installed** in Microsoft Word
- **Document session created** with `document_id`

### Example: Live Search & Replace
```python
# Via Claude Code or MCP client
await enhanced_search_and_replace(
    document_id="main",  # Must use document_id for live editing
    find_text="important",
    replace_text="CRITICAL",
    apply_formatting=True,
    bold=True,
    color="red"
)
```

### Example: Live Content Analysis
```python
# Extract content from live document
content = await get_text(
    document_id="main",
    scope="search", 
    search_term="methodology"
)
```

### Document Session Setup
```python
# First, create a document session
session_manager = get_session_manager()
session_manager.open_document("main", "/path/to/document.docx")

# Then the Add-in can connect to this session
# The Add-in uses the document's file path to find the right session
```

---

## Troubleshooting

### 🔧 Common Issues

**Add-in shows "SEARCHING MCP..."**
- ✅ Make sure you ran the startup script
- ✅ Check that port 8765 is not blocked
- ✅ Verify MCP server started successfully

**Connection fails repeatedly**
- ✅ Restart both servers using the startup script
- ✅ Check Windows Firewall/macOS firewall settings
- ✅ Try manually clicking "Connect" button

**"No live session found" error**
- ✅ Create a document session first: `session_manager.open_document("id", "path")`
- ✅ Make sure document path matches what Add-in sends
- ✅ Use `document_id` parameter, not `filename`

### 🔍 Debug Information

**Check server status:**
```bash
# Check if MCP server is running
lsof -i :8765

# Check if Add-in server is running  
lsof -i :3000
```

**Console logs:**
- **Word Add-in**: Open Developer Tools in Word (F12)
- **MCP Server**: Check terminal output for WebSocket messages

---

## Manual Setup (Alternative)

If you prefer manual control:

### 1. Start MCP Server
```bash
python -m word_document_server.main
```

### 2. Start Add-in Server
```bash
cd word-live-addin
npm install  # First time only
npm start
```

### 3. Load Add-in
- Insert → My Add-ins → Upload → `manifest.xml`

---

## Advanced Configuration

### Port Configuration
- **MCP Server**: Edit `WEBSOCKET_URL` in `taskpane.js` (default: 8765)
- **Add-in Server**: Edit `package.json` scripts (default: 3000)

### Auto-Connect Settings
```javascript
// In taskpane.js
const MAX_RETRY_ATTEMPTS = 10;  // Max connection attempts
const RETRY_INTERVAL = 2000;    // 2 seconds between attempts
```

### Connection Timeout
```python
# In session_manager.py
await asyncio.wait_for(future, timeout=60.0)  # 60 second timeout
```

---

## 🎉 Success Indicators

✅ **MCP Server**: Terminal shows "WebSocket server started on ws://localhost:8765"  
✅ **Add-in Server**: Browser opens to https://localhost:3000  
✅ **Add-in Loaded**: Task pane appears in Word  
✅ **Connected**: Status shows green "LIVE" indicator  
✅ **Working**: Live editing commands execute in real-time  

**When everything is working, you'll see:**
- 💚 Green "LIVE" status in Word task pane
- 📡 WebSocket messages in server terminal
- ⚡ Instant document updates from AI commands

---

## Next Steps

1. **Try live editing** with Claude Code using `document_id` parameters
2. **Create document sessions** for multiple documents
3. **Experiment** with real-time search & replace, content extraction
4. **Build custom workflows** combining multiple live editing operations

**Happy live editing!** 🎯