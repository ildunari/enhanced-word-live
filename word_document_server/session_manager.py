"""
Document Session Manager for Word Document Server.

Provides in-memory storage and management of open Word documents with simple IDs,
eliminating the need to pass full file paths for every operation.
"""
import os
import asyncio
import json
import uuid
from typing import Dict, Optional, List, Any
from dataclasses import dataclass, field
from docx import Document
from word_document_server.utils.file_utils import ensure_docx_extension


@dataclass
class DocumentHandle:
    """Container for an open Word document with metadata and optional live connection."""
    document_id: str
    file_path: str
    document: Document
    metadata: Dict[str, Any]
    websocket_connection: Optional[object] = None
    pending_requests: Dict[str, asyncio.Future] = field(default_factory=dict)
    
    def __post_init__(self):
        """Ensure file path has .docx extension."""
        self.file_path = ensure_docx_extension(self.file_path)
    
    @property
    def is_live(self) -> bool:
        """Check if this document has an active live WebSocket connection."""
        return self.websocket_connection is not None
    
    def register_websocket(self, websocket: object):
        """Register a WebSocket connection for live editing."""
        self.websocket_connection = websocket
        print(f"[DocumentHandle] Registered live connection for {self.document_id}")
    
    def unregister_websocket(self):
        """Remove the WebSocket connection."""
        if self.websocket_connection:
            print(f"[DocumentHandle] Unregistered live connection for {self.document_id}")
            self.websocket_connection = None
            # Cancel any pending requests
            for future in self.pending_requests.values():
                if not future.done():
                    future.cancel()
            self.pending_requests.clear()



class DocumentSessionManager:
    """Manages open Word documents with simple string IDs."""
    
    def __init__(self):
        self._documents: Dict[str, DocumentHandle] = {}
        self._active_document_id: Optional[str] = None
    
    def open_document(self, document_id: str, file_path: str) -> str:
        """
        Open a Word document and assign it a simple ID for future operations.
        
        Args:
            document_id: Simple identifier for the document (e.g., "main", "draft", "review")
            file_path: Full path to the Word document file
            
        Returns:
            Success/error message string
        """
        try:
            # Validate inputs
            if not document_id or not document_id.strip():
                return "Error: document_id cannot be empty"
                
            if not file_path or not file_path.strip():
                return "Error: file_path cannot be empty"
            
            # Ensure .docx extension
            file_path = ensure_docx_extension(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return f"Error: Document file '{file_path}' does not exist"
            
            # Check if document_id already in use
            if document_id in self._documents:
                return f"Error: Document ID '{document_id}' is already in use. Use close_document() first or choose a different ID."
            
            # Try to open the document
            try:
                doc = Document(file_path)
            except Exception as e:
                return f"Error: Failed to open document '{file_path}': {str(e)}"
            
            # Create document handle with metadata
            metadata = {
                "opened_at": str(os.path.getmtime(file_path)),
                "paragraph_count": len(doc.paragraphs),
                "section_count": len(doc.sections),
                "file_size": os.path.getsize(file_path)
            }
            
            handle = DocumentHandle(
                document_id=document_id,
                file_path=file_path,
                document=doc,
                metadata=metadata
            )
            
            # Store in session
            self._documents[document_id] = handle
            
            # Set as active if it's the first document
            if self._active_document_id is None:
                self._active_document_id = document_id
                
            return f"Successfully opened document '{document_id}' from '{file_path}'"
            
        except Exception as e:
            return f"Error opening document: {str(e)}"
    
    def close_document(self, document_id: str) -> str:
        """
        Close a document and remove it from the session.
        
        Args:
            document_id: ID of the document to close
            
        Returns:
            Success/error message string
        """
        try:
            if document_id not in self._documents:
                return f"Error: Document ID '{document_id}' not found in session"
            
            # Get handle before removing
            handle = self._documents[document_id]
            
            # Remove from session
            del self._documents[document_id]
            
            # Update active document if needed
            if self._active_document_id == document_id:
                # Set active to another document if available, or None
                self._active_document_id = next(iter(self._documents.keys())) if self._documents else None
            
            return f"Successfully closed document '{document_id}' (was '{handle.file_path}')"
            
        except Exception as e:
            return f"Error closing document: {str(e)}"
    
    def list_open_documents(self) -> str:
        """
        List all currently open documents with their metadata.
        
        Returns:
            Formatted string with document information
        """
        try:
            if not self._documents:
                return "No documents currently open"
            
            result = f"Open documents ({len(self._documents)}):\n\n"
            
            for doc_id, handle in self._documents.items():
                is_active = " (ACTIVE)" if doc_id == self._active_document_id else ""
                result += f"ID: {doc_id}{is_active}\n"
                result += f"  Path: {handle.file_path}\n"
                result += f"  Paragraphs: {handle.metadata.get('paragraph_count', 'Unknown')}\n"
                result += f"  Sections: {handle.metadata.get('section_count', 'Unknown')}\n"
                result += f"  File size: {handle.metadata.get('file_size', 'Unknown')} bytes\n\n"
            
            return result.rstrip()
            
        except Exception as e:
            return f"Error listing documents: {str(e)}"
    
    def set_active_document(self, document_id: str) -> str:
        """
        Set the active/default document for operations that support it.
        
        Args:
            document_id: ID of the document to make active
            
        Returns:
            Success/error message string
        """
        try:
            if document_id not in self._documents:
                return f"Error: Document ID '{document_id}' not found in session"
            
            old_active = self._active_document_id
            self._active_document_id = document_id
            
            if old_active:
                return f"Active document changed from '{old_active}' to '{document_id}'"
            else:
                return f"Active document set to '{document_id}'"
                
        except Exception as e:
            return f"Error setting active document: {str(e)}"
    
    def get_document(self, document_id: str) -> Optional[DocumentHandle]:
        """
        Get a document handle by ID.
        
        Args:
            document_id: ID of the document to retrieve
            
        Returns:
            DocumentHandle if found, None otherwise
        """
        return self._documents.get(document_id)
    
    def get_document_path(self, document_id: str) -> Optional[str]:
        """
        Get the file path for a document by ID.
        
        Args:
            document_id: ID of the document
            
        Returns:
            File path if document found, None otherwise
        """
        handle = self.get_document(document_id)
        return handle.file_path if handle else None
    
    def validate_document_id(self, document_id: str) -> str:
        """
        Validate that a document ID exists in the session.
        
        Args:
            document_id: ID to validate
            
        Returns:
            Empty string if valid, error message if invalid
        """
        if not document_id:
            return "Error: document_id parameter is required"
        
        if document_id not in self._documents:
            available = list(self._documents.keys())
            if available:
                return f"Error: Document ID '{document_id}' not found. Available: {', '.join(available)}"
            else:
                return f"Error: Document ID '{document_id}' not found. No documents are currently open. Use open_document() first."
        
        return ""  # Valid
    
    def close_all_documents(self) -> str:
        """
        Close all open documents.
        
        Returns:
            Success message with count
        """
        count = len(self._documents)
        self._documents.clear()
        self._active_document_id = None
        return f"Closed {count} documents"
    
    # Live Session Management Methods
    
    def register_live_connection(self, document_id: str, websocket: object) -> str:
        """
        Register a WebSocket connection for live editing of a document.
        
        Args:
            document_id: ID of the document to make live
            websocket: WebSocket connection object
            
        Returns:
            Success/error message string
        """
        try:
            handle = self.get_document(document_id)
            if not handle:
                return f"Error: Document ID '{document_id}' not found in session"
            
            handle.register_websocket(websocket)
            return f"Successfully registered live connection for document '{document_id}'"
            
        except Exception as e:
            return f"Error registering live connection: {str(e)}"
        
    def unregister_live_connection(self, document_id: str) -> str:
        """
        Remove WebSocket connection from a document.
        
        Args:
            document_id: ID of the document to disconnect
            
        Returns:
            Success/error message string
        """
        try:
            handle = self.get_document(document_id)
            if not handle:
                return f"Error: Document ID '{document_id}' not found in session"
            
            handle.unregister_websocket()
            return f"Successfully unregistered live connection for document '{document_id}'"
            
        except Exception as e:
            return f"Error unregistering live connection: {str(e)}"
        
    def find_document_by_websocket(self, websocket: object) -> Optional[str]:
        """
        Find the document ID associated with a WebSocket connection.
        
        Args:
            websocket: WebSocket connection to search for
            
        Returns:
            Document ID if found, None otherwise
        """
        for doc_id, handle in self._documents.items():
            if handle.websocket_connection == websocket:
                return doc_id
        return None
    
    def is_document_live(self, document_id: str) -> bool:
        """
        Check if a document has an active live connection.
        
        Args:
            document_id: ID of the document to check
            
        Returns:
            True if document is live, False otherwise
        """
        handle = self.get_document(document_id)
        return handle.is_live if handle else False
        
    async def send_live_request(self, document_id: str, command: str, **kwargs) -> dict:
        """
        Send a request to the live Word Add-in via WebSocket.
        
        Args:
            document_id: ID of the document
            command: Command to send to Add-in
            **kwargs: Additional parameters for the command
            
        Returns:
            Response data from Add-in
            
        Raises:
            ConnectionError: If no live session found
            TimeoutError: If request times out
        """
        handle = self.get_document(document_id)
        if not handle or not handle.is_live:
            raise ConnectionError(f"No live session found for document: {document_id}")
        
        correlation_id = str(uuid.uuid4())
        future = asyncio.get_running_loop().create_future()
        handle.pending_requests[correlation_id] = future
        
        request_message = {
            "command": command,
            "correlation_id": correlation_id,
            **kwargs
        }
        
        try:
            print(f"[DocumentSessionManager] Sending command '{command}' to Add-in for document '{document_id}'")
            # WebSocket send requires JSON string serialization
            await handle.websocket_connection.send(json.dumps(request_message))
            return await asyncio.wait_for(future, timeout=60.0)
        except asyncio.TimeoutError:
            handle.pending_requests.pop(correlation_id, None)
            raise TimeoutError(f"Request to Add-in for command '{command}' timed out.")
        except Exception as e:
            handle.pending_requests.pop(correlation_id, None)
            raise e
        
    def handle_live_response(self, websocket: object, correlation_id: str, data: dict):
        """
        Handle a response from the Word Add-in.
        
        Args:
            websocket: WebSocket connection that sent the response
            correlation_id: ID of the request being responded to
            data: Response data from Add-in
        """
        document_id = self.find_document_by_websocket(websocket)
        if not document_id:
            print(f"[DocumentSessionManager] WARN: Received response from unknown websocket")
            return
        
        handle = self.get_document(document_id)
        if not handle:
            return
        
        if correlation_id in handle.pending_requests:
            future = handle.pending_requests.pop(correlation_id)
            if data.get("status") == "success":
                future.set_result(data.get("data"))
            else:
                future.set_exception(Exception(data.get("error", "Unknown error from Add-in")))
        else:
            print(f"[DocumentSessionManager] WARN: Received response for unknown correlation_id: {correlation_id}")

# Global session manager instance
_session_manager = DocumentSessionManager()


def get_session_manager() -> DocumentSessionManager:
    """Get the global document session manager instance."""
    return _session_manager