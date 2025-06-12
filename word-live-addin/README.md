# Word Live Connector Add-in

This Office Add-in connects a live Word document to the Enhanced Word MCP Server for real-time editing capabilities.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```

2. Start the development server:
   ```bash
   npm start
   ```

3. Load the add-in in Word:
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select the `manifest.xml` file from this directory

## Usage

1. Open a Word document
2. Open the Word Live Connector task pane
3. Click "Connect" to establish a live connection
4. The Enhanced Word MCP Server can now edit the document in real-time

## Requirements

- Microsoft Word (Desktop version)
- Node.js and npm
- Enhanced Word MCP Server running with WebSocket support