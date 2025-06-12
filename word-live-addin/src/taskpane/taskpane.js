/*
 * Word Live Connector: Bridge between a live Word document and the local MCP server.
 * This script manages the WebSocket connection and executes commands via the Office.js API.
 */

/* global document, Office, Word */

const WEBSOCKET_URL = "ws://localhost:8765";
let socket = null;
let documentPath = ""; // Will hold the unique URL or generated ID for this document.
let connectButton;
let statusText;
let statusIndicator;
let autoConnectInterval = null;
let retryCount = 0;
const MAX_RETRY_ATTEMPTS = 10;
const RETRY_INTERVAL = 2000; // 2 seconds

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // DOM elements are ready, assign them.
    connectButton = document.getElementById("connect-button");
    statusText = document.getElementById("status-text");
    statusIndicator = document.getElementById("status-indicator");
    
    // Assign the primary event handler.
    connectButton.onclick = toggleConnection;
    updateUiState("disconnected");
    
    // Start automatic connection detection
    startAutoConnectDetection();
  }
});

function updateUiState(state, message = "") {
  switch (state) {
    case "disconnected":
      statusIndicator.className = "status-indicator-gray";
      statusText.textContent = "DISCONNECTED";
      connectButton.textContent = "Connect";
      connectButton.classList.remove("is-connected");
      connectButton.disabled = false;
      break;
    case "searching":
      statusIndicator.className = "status-indicator-orange";
      statusText.textContent = "SEARCHING MCP...";
      connectButton.textContent = "Searching...";
      connectButton.disabled = true;
      break;
    case "connecting":
      statusIndicator.className = "status-indicator-orange";
      statusText.textContent = "CONNECTING...";
      connectButton.textContent = "Connecting...";
      connectButton.disabled = true;
      break;
    case "connected":
      statusIndicator.className = "status-indicator-green";
      statusText.textContent = "LIVE";
      connectButton.textContent = "Disconnect";
      connectButton.classList.add("is-connected");
      connectButton.disabled = false;
      // Stop auto-connect detection when successfully connected
      stopAutoConnectDetection();
      break;
    case "error":
      statusIndicator.className = "status-indicator-red";
      statusText.textContent = message || "ERROR";
      connectButton.textContent = "Connect";
      connectButton.classList.remove("is-connected");
      connectButton.disabled = false;
      break;
  }
}

function toggleConnection() {
  if (socket && socket.readyState === WebSocket.OPEN) {
    disconnect();
  } else {
    // Stop auto-detection and connect manually
    stopAutoConnectDetection();
    connect();
  }
}

function startAutoConnectDetection() {
  console.log("Starting automatic MCP server detection...");
  updateUiState("searching");
  
  autoConnectInterval = setInterval(() => {
    if (socket && socket.readyState === WebSocket.OPEN) {
      // Already connected, stop detection
      stopAutoConnectDetection();
      return;
    }
    
    console.log(`Auto-connect attempt ${retryCount + 1}/${MAX_RETRY_ATTEMPTS}`);
    attemptAutoConnect();
  }, RETRY_INTERVAL);
  
  // Try connecting immediately
  attemptAutoConnect();
}

function stopAutoConnectDetection() {
  if (autoConnectInterval) {
    console.log("Stopping automatic MCP server detection");
    clearInterval(autoConnectInterval);
    autoConnectInterval = null;
    retryCount = 0;
  }
}

function attemptAutoConnect() {
  if (retryCount >= MAX_RETRY_ATTEMPTS) {
    console.log("Max auto-connect attempts reached. Stopping detection.");
    stopAutoConnectDetection();
    updateUiState("disconnected");
    return;
  }
  
  retryCount++;
  
  // Test if WebSocket server is available
  const testSocket = new WebSocket(WEBSOCKET_URL);
  
  testSocket.onopen = () => {
    console.log("MCP server detected! Establishing connection...");
    testSocket.close();
    connect();
  };
  
  testSocket.onerror = () => {
    console.log(`Auto-connect attempt ${retryCount} failed - MCP server not available`);
    // Continue trying in the next interval
  };
  
  testSocket.onclose = () => {
    // Normal close after testing
  };
}

function connect() {
  updateUiState("connecting");
  
  // Asynchronously get the file's URL to use as a unique session key.
  Office.context.document.getFilePropertiesAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Error getting document properties: " + asyncResult.error.message);
      updateUiState("error", "SETUP FAILED");
      return;
    }
    
    documentPath = asyncResult.value.url;
    // Fallback for environments where the URL is not available (e.g., new unsaved docs).
    if (!documentPath || documentPath.trim() === "") {
      documentPath = "live-unsaved-doc-" + Date.now();
      console.warn("Document URL not available. Using generated ID:", documentPath);
    }
    
    console.log(`Attempting WebSocket connection for document: ${documentPath}`);
    
    socket = new WebSocket(WEBSOCKET_URL);

    socket.onopen = (event) => {
      console.log("WebSocket connection established.");
      updateUiState("connected");
      
      const registrationMessage = {
        type: "register",
        path: documentPath,
      };
      socket.send(JSON.stringify(registrationMessage));
    };

    socket.onmessage = (event) => {
      handleServerMessage(JSON.parse(event.data));
    };

    socket.onclose = (event) => {
      console.log("WebSocket connection closed.", event.code, event.reason);
      updateUiState("disconnected");
      socket = null;
      
      // Restart auto-detection if connection was lost unexpectedly
      if (event.code !== 1000) { // 1000 = normal closure
        console.log("Connection lost unexpectedly, restarting auto-detection...");
        setTimeout(startAutoConnectDetection, 3000); // Wait 3 seconds before retrying
      }
    };

    socket.onerror = (error) => {
      console.error("WebSocket Error:", error);
      updateUiState("error", "CONNECTION FAILED");
      socket = null;
    };
  });
}

function disconnect() {
  // Stop auto-detection when manually disconnecting
  stopAutoConnectDetection();
  
  if (socket) {
    socket.close(1000, "Manual disconnect"); // 1000 = normal closure
  }
}

async function handleServerMessage(message) {
  console.log("Message from server: ", message);
  const { command, correlation_id } = message;

  try {
    let result_data = {};
    let command_status = "success";
    
    switch (command) {
      case "get_full_content":
        result_data = { content: await getDocumentAsOoxml() };
        break;
      
      case "replace_full_content":
        await setDocumentFromOoxml(message.content);
        break;

      case "find_and_format":
        const count = await findAndFormat(message.find_text, message.formatting);
        result_data = { replacements: count };
        break;

      default:
        command_status = "error";
        result_data = { error: `Unknown command: ${command}` };
        throw new Error(result_data.error);
    }
    
    socket.send(JSON.stringify({ type: "response", status: command_status, correlation_id, data: result_data }));

  } catch (error) {
    console.error(`Error processing command '${command}':`, error);
    socket.send(JSON.stringify({ type: "response", status: "error", correlation_id, error: error.message }));
  }
}

// --- Office.js API Wrapper Functions ---

function getDocumentAsOoxml() {
  return Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    return ooxml.value;
  });
}

function setDocumentFromOoxml(ooxml) {
    return Word.run(async (context) => {
        context.document.body.clear();
        context.document.body.insertOoxml(ooxml, Word.InsertLocation.replace);
        await context.sync();
    });
}

function findAndFormat(searchText, formatting) {
  return Word.run(async (context) => {
    // This is an optimized path that uses Word's native search.
    const searchResults = context.document.body.search(searchText, { matchCase: formatting.matchCase || false });
    context.load(searchResults, 'font');
    await context.sync();

    searchResults.items.forEach(item => {
      if (formatting.bold !== undefined) item.font.bold = formatting.bold;
      if (formatting.italic !== undefined) item.font.italic = formatting.italic;
      if (formatting.underline !== undefined) item.font.underline = formatting.underline;
      if (formatting.color !== undefined) item.font.color = formatting.color;
      if (formatting.font_name !== undefined) item.font.name = formatting.font_name;
      if (formatting.font_size !== undefined) item.font.size = formatting.font_size;
    });

    await context.sync();
    return searchResults.items.length;
  });
}