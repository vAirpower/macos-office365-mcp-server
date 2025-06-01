#!/usr/bin/env python3
"""
Office 365 MCP Server for macOS
A comprehensive MCP server for Microsoft PowerPoint and Word automation on macOS.
"""

import asyncio
import logging
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# MCP imports
try:
    from mcp.server import FastMCP
except ImportError:
    # Fallback for different MCP versions
    try:
        from mcp import Server as FastMCP
    except ImportError:
        # Create a basic server implementation
        class FastMCP:
            def __init__(self, name):
                self.name = name
                self.tools = {}
                self.resources = {}
            
            def tool(self):
                def decorator(func):
                    self.tools[func.__name__] = func
                    return func
                return decorator
            
            def resource(self, uri):
                def decorator(func):
                    self.resources[uri] = func
                    return func
                return decorator
            
            def run(self):
                print(f"Starting {self.name}...")
                print(f"Available tools: {list(self.tools.keys())}")
                print(f"Available resources: {list(self.resources.keys())}")
                # Basic server loop
                import time
                try:
                    while True:
                        time.sleep(1)
                except KeyboardInterrupt:
                    print("Server stopped.")

# Local imports
from controllers.powerpoint_controller import PowerPointController
from controllers.word_controller import WordController
from integrations.applescript_bridge import AppleScriptBridge
from utils.config import Config
from utils.logger import setup_logger
from utils.validators import validate_input

# Initialize logging
logger = setup_logger(__name__)

# Initialize FastMCP server
mcp = FastMCP("Office365 MCP Server")

# Initialize controllers
powerpoint = PowerPointController()
word = WordController()
applescript = AppleScriptBridge()
config = Config()

# Track active documents/presentations
active_presentations: Dict[str, Any] = {}
active_documents: Dict[str, Any] = {}

# PowerPoint Tools
@mcp.tool()
async def create_presentation(
    title: str,
    theme: str = "default",
    template_path: Optional[str] = None
) -> Dict[str, Any]:
    """Create a new PowerPoint presentation.
    
    Args:
        title: Presentation title
        theme: Theme name (default, modern, classic, etc.)
        template_path: Optional path to custom template
        
    Returns:
        Dict with presentation_id and metadata
    """
    try:
        result = await powerpoint.create_presentation(
            title=title,
            theme=theme,
            template_path=template_path
        )
        
        # Store in active presentations
        active_presentations[result["presentation_id"]] = result
        
        logger.info(f"Created presentation: {title}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to create presentation: {e}")
        raise

@mcp.tool()
async def add_slide(
    presentation_id: str,
    layout: str = "Title and Content",
    position: Optional[int] = None
) -> Dict[str, Any]:
    """Add a new slide to a presentation.
    
    Args:
        presentation_id: ID of the presentation
        layout: Slide layout name
        position: Position to insert slide (None for end)
        
    Returns:
        Dict with slide_id and metadata
    """
    try:
        result = await powerpoint.add_slide(
            presentation_id=presentation_id,
            layout=layout,
            position=position
        )
        
        logger.info(f"Added slide to presentation {presentation_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add slide: {e}")
        raise

@mcp.tool()
async def add_text_to_slide(
    slide_id: str,
    text: str,
    placeholder: str = "content",
    formatting: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """Add text content to a slide.
    
    Args:
        slide_id: ID of the slide
        text: Text content to add
        placeholder: Placeholder name (title, content, etc.)
        formatting: Text formatting options
        
    Returns:
        Dict with operation status
    """
    try:
        result = await powerpoint.add_text(
            slide_id=slide_id,
            text=text,
            placeholder=placeholder,
            formatting=formatting or {}
        )
        
        logger.info(f"Added text to slide {slide_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add text: {e}")
        raise

@mcp.tool()
async def add_image_to_slide(
    slide_id: str,
    image_source: str,
    position: Optional[Dict[str, float]] = None,
    size: Optional[Dict[str, float]] = None
) -> Dict[str, Any]:
    """Add an image to a slide.
    
    Args:
        slide_id: ID of the slide
        image_source: Path to image file or URL
        position: Position dict with x, y coordinates
        size: Size dict with width, height
        
    Returns:
        Dict with operation status
    """
    try:
        result = await powerpoint.add_image(
            slide_id=slide_id,
            image_source=image_source,
            position=position or {"x": 100, "y": 100},
            size=size or {"width": 400, "height": 300}
        )
        
        logger.info(f"Added image to slide {slide_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add image: {e}")
        raise

@mcp.tool()
async def add_speaker_notes(
    slide_id: str,
    notes: str
) -> Dict[str, Any]:
    """Add speaker notes to a slide.
    
    Args:
        slide_id: ID of the slide
        notes: Speaker notes content
        
    Returns:
        Dict with operation status
    """
    try:
        result = await powerpoint.add_speaker_notes(
            slide_id=slide_id,
            notes=notes
        )
        
        logger.info(f"Added speaker notes to slide {slide_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add speaker notes: {e}")
        raise

@mcp.tool()
async def save_presentation(
    presentation_id: str,
    file_path: str,
    format: str = "pptx"
) -> Dict[str, Any]:
    """Save a presentation to file.
    
    Args:
        presentation_id: ID of the presentation
        file_path: Path to save the file
        format: File format (pptx, pdf, etc.)
        
    Returns:
        Dict with operation status and file path
    """
    try:
        result = await powerpoint.save_presentation(
            presentation_id=presentation_id,
            file_path=file_path,
            format=format
        )
        
        logger.info(f"Saved presentation to {file_path}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to save presentation: {e}")
        raise

# Word Tools
@mcp.tool()
async def create_document(
    title: str = "New Document",
    template_path: Optional[str] = None
) -> Dict[str, Any]:
    """Create a new Word document.
    
    Args:
        title: Document title
        template_path: Optional path to template
        
    Returns:
        Dict with document_id and metadata
    """
    try:
        result = await word.create_document(
            title=title,
            template_path=template_path
        )
        
        # Store in active documents
        active_documents[result["document_id"]] = result
        
        logger.info(f"Created document: {title}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to create document: {e}")
        raise

@mcp.tool()
async def add_heading(
    document_id: str,
    text: str,
    level: int = 1,
    style: Optional[str] = None
) -> Dict[str, Any]:
    """Add a heading to a document.
    
    Args:
        document_id: ID of the document
        text: Heading text
        level: Heading level (1-6)
        style: Optional style name
        
    Returns:
        Dict with operation status
    """
    try:
        result = await word.add_heading(
            document_id=document_id,
            text=text,
            level=level,
            style=style
        )
        
        logger.info(f"Added heading to document {document_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add heading: {e}")
        raise

@mcp.tool()
async def add_paragraph(
    document_id: str,
    text: str,
    style: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """Add a paragraph to a document.
    
    Args:
        document_id: ID of the document
        text: Paragraph text
        style: Optional style name
        formatting: Text formatting options
        
    Returns:
        Dict with operation status
    """
    try:
        result = await word.add_paragraph(
            document_id=document_id,
            text=text,
            style=style,
            formatting=formatting or {}
        )
        
        logger.info(f"Added paragraph to document {document_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add paragraph: {e}")
        raise

@mcp.tool()
async def add_list(
    document_id: str,
    items: List[str],
    list_type: str = "bullet",
    style: Optional[str] = None
) -> Dict[str, Any]:
    """Add a list to a document.
    
    Args:
        document_id: ID of the document
        items: List items
        list_type: Type of list (bullet, number)
        style: Optional style name
        
    Returns:
        Dict with operation status
    """
    try:
        result = await word.add_list(
            document_id=document_id,
            items=items,
            list_type=list_type,
            style=style
        )
        
        logger.info(f"Added list to document {document_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add list: {e}")
        raise

@mcp.tool()
async def add_table(
    document_id: str,
    rows: int,
    columns: int,
    data: Optional[List[List[str]]] = None,
    style: Optional[str] = None
) -> Dict[str, Any]:
    """Add a table to a document.
    
    Args:
        document_id: ID of the document
        rows: Number of rows
        columns: Number of columns
        data: Optional table data
        style: Optional table style
        
    Returns:
        Dict with operation status
    """
    try:
        result = await word.add_table(
            document_id=document_id,
            rows=rows,
            columns=columns,
            data=data,
            style=style
        )
        
        logger.info(f"Added table to document {document_id}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to add table: {e}")
        raise

@mcp.tool()
async def save_document(
    document_id: str,
    file_path: str,
    format: str = "docx"
) -> Dict[str, Any]:
    """Save a document to file.
    
    Args:
        document_id: ID of the document
        file_path: Path to save the file
        format: File format (docx, pdf, etc.)
        
    Returns:
        Dict with operation status and file path
    """
    try:
        result = await word.save_document(
            document_id=document_id,
            file_path=file_path,
            format=format
        )
        
        logger.info(f"Saved document to {file_path}")
        return result
        
    except Exception as e:
        logger.error(f"Failed to save document: {e}")
        raise

# Utility Tools
@mcp.tool()
async def list_active_presentations() -> List[Dict[str, Any]]:
    """List all active presentations.
    
    Returns:
        List of active presentation metadata
    """
    return list(active_presentations.values())

@mcp.tool()
async def list_active_documents() -> List[Dict[str, Any]]:
    """List all active documents.
    
    Returns:
        List of active document metadata
    """
    return list(active_documents.values())

@mcp.tool()
async def check_office_status() -> Dict[str, Any]:
    """Check if Office applications are available.
    
    Returns:
        Dict with Office application status
    """
    try:
        powerpoint_status = await applescript.check_powerpoint_status()
        word_status = await applescript.check_word_status()
        
        return {
            "powerpoint_available": powerpoint_status,
            "word_available": word_status,
            "server_status": "running"
        }
        
    except Exception as e:
        logger.error(f"Failed to check Office status: {e}")
        return {
            "powerpoint_available": False,
            "word_available": False,
            "server_status": "error",
            "error": str(e)
        }

# Resources
@mcp.resource("office365://templates")
async def get_templates() -> str:
    """Get available Office templates."""
    templates_dir = Path(__file__).parent / "templates"
    templates = []
    
    if templates_dir.exists():
        for template_file in templates_dir.glob("*"):
            templates.append({
                "name": template_file.stem,
                "path": str(template_file),
                "type": template_file.suffix[1:]
            })
    
    return f"Available templates: {templates}"

@mcp.resource("office365://status")
async def get_server_status() -> str:
    """Get server status information."""
    status = {
        "active_presentations": len(active_presentations),
        "active_documents": len(active_documents),
        "server_version": "1.0.0",
        "platform": "macOS"
    }
    
    return f"Server Status: {status}"

if __name__ == "__main__":
    logger.info("Starting Office 365 MCP Server...")
    mcp.run()
