"""
AppleScript Bridge for Office 365 MCP Server
Provides direct integration with Microsoft Office applications on macOS via AppleScript.
"""

import asyncio
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import logging

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils.logger import setup_logger

logger = setup_logger(__name__)

class AppleScriptBridge:
    """Bridge for communicating with Office applications via AppleScript."""
    
    def __init__(self):
        self.powerpoint_app = "Microsoft PowerPoint"
        self.word_app = "Microsoft Word"
    
    async def execute_applescript(self, script: str) -> str:
        """Execute an AppleScript and return the result.
        
        Args:
            script: AppleScript code to execute
            
        Returns:
            Script output as string
        """
        try:
            # Use osascript to execute AppleScript
            process = await asyncio.create_subprocess_exec(
                "osascript", "-e", script,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            stdout, stderr = await process.communicate()
            
            if process.returncode != 0:
                error_msg = stderr.decode().strip()
                logger.error(f"AppleScript error: {error_msg}")
                raise RuntimeError(f"AppleScript execution failed: {error_msg}")
            
            return stdout.decode().strip()
            
        except Exception as e:
            logger.error(f"Failed to execute AppleScript: {e}")
            raise
    
    async def check_powerpoint_status(self) -> bool:
        """Check if PowerPoint is available and running.
        
        Returns:
            True if PowerPoint is available
        """
        try:
            script = f'''
            tell application "System Events"
                return (name of processes) contains "{self.powerpoint_app}"
            end tell
            '''
            
            result = await self.execute_applescript(script)
            return result.lower() == "true"
            
        except Exception as e:
            logger.warning(f"Could not check PowerPoint status: {e}")
            return False
    
    async def check_word_status(self) -> bool:
        """Check if Word is available and running.
        
        Returns:
            True if Word is available
        """
        try:
            script = f'''
            tell application "System Events"
                return (name of processes) contains "{self.word_app}"
            end tell
            '''
            
            result = await self.execute_applescript(script)
            return result.lower() == "true"
            
        except Exception as e:
            logger.warning(f"Could not check Word status: {e}")
            return False
    
    async def launch_powerpoint(self) -> bool:
        """Launch PowerPoint application.
        
        Returns:
            True if successfully launched
        """
        try:
            script = f'''
            tell application "{self.powerpoint_app}"
                activate
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info("PowerPoint launched successfully")
            return True
            
        except Exception as e:
            logger.error(f"Failed to launch PowerPoint: {e}")
            return False
    
    async def launch_word(self) -> bool:
        """Launch Word application.
        
        Returns:
            True if successfully launched
        """
        try:
            script = f'''
            tell application "{self.word_app}"
                activate
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info("Word launched successfully")
            return True
            
        except Exception as e:
            logger.error(f"Failed to launch Word: {e}")
            return False
    
    async def open_powerpoint_file(self, file_path: str) -> bool:
        """Open a PowerPoint file.
        
        Args:
            file_path: Path to the PowerPoint file
            
        Returns:
            True if successfully opened
        """
        try:
            # Ensure PowerPoint is running
            await self.launch_powerpoint()
            
            script = f'''
            tell application "{self.powerpoint_app}"
                open POSIX file "{file_path}"
                activate
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Opened PowerPoint file: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to open PowerPoint file: {e}")
            return False
    
    async def open_word_file(self, file_path: str) -> bool:
        """Open a Word file.
        
        Args:
            file_path: Path to the Word file
            
        Returns:
            True if successfully opened
        """
        try:
            # Ensure Word is running
            await self.launch_word()
            
            script = f'''
            tell application "{self.word_app}"
                open POSIX file "{file_path}"
                activate
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Opened Word file: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to open Word file: {e}")
            return False
    
    async def create_powerpoint_presentation(self, title: str = "New Presentation") -> Dict[str, Any]:
        """Create a new PowerPoint presentation via AppleScript.
        
        Args:
            title: Presentation title
            
        Returns:
            Dict with presentation information
        """
        try:
            await self.launch_powerpoint()
            
            script = f'''
            tell application "{self.powerpoint_app}"
                set newPres to make new presentation
                tell newPres
                    set slide1 to make new slide at beginning
                    tell slide1
                        set title of text range of text frame of shape 1 to "{title}"
                    end tell
                end tell
                return name of newPres
            end tell
            '''
            
            result = await self.execute_applescript(script)
            logger.info(f"Created PowerPoint presentation via AppleScript: {title}")
            
            return {
                "status": "success",
                "title": title,
                "applescript_name": result,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to create PowerPoint presentation via AppleScript: {e}")
            raise
    
    async def add_slide_to_presentation(self, presentation_name: str, layout: str = "Title and Content") -> Dict[str, Any]:
        """Add a slide to a PowerPoint presentation via AppleScript.
        
        Args:
            presentation_name: Name of the presentation
            layout: Slide layout
            
        Returns:
            Dict with slide information
        """
        try:
            # Map layout names to PowerPoint constants
            layout_map = {
                "Title Slide": "title slide",
                "Title and Content": "title and content",
                "Section Header": "section header",
                "Two Content": "two content",
                "Comparison": "comparison",
                "Title Only": "title only",
                "Blank": "blank",
                "Content with Caption": "content with caption",
                "Picture with Caption": "picture with caption"
            }
            
            ppt_layout = layout_map.get(layout, "title and content")
            
            script = f'''
            tell application "{self.powerpoint_app}"
                tell presentation "{presentation_name}"
                    set newSlide to make new slide at end
                    return index of newSlide
                end tell
            end tell
            '''
            
            result = await self.execute_applescript(script)
            slide_index = int(result)
            
            logger.info(f"Added slide to presentation {presentation_name}")
            
            return {
                "status": "success",
                "presentation_name": presentation_name,
                "slide_index": slide_index,
                "layout": layout,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to add slide via AppleScript: {e}")
            raise
    
    async def add_text_to_slide(
        self,
        presentation_name: str,
        slide_index: int,
        text: str,
        placeholder: str = "content"
    ) -> Dict[str, Any]:
        """Add text to a slide via AppleScript.
        
        Args:
            presentation_name: Name of the presentation
            slide_index: Index of the slide
            text: Text to add
            placeholder: Placeholder type
            
        Returns:
            Dict with operation status
        """
        try:
            # Escape quotes in text
            escaped_text = text.replace('"', '\\"')
            
            if placeholder == "title":
                shape_index = 1  # Title is usually shape 1
            else:
                shape_index = 2  # Content is usually shape 2
            
            script = f'''
            tell application "{self.powerpoint_app}"
                tell presentation "{presentation_name}"
                    tell slide {slide_index}
                        set text range of text frame of shape {shape_index} to "{escaped_text}"
                    end tell
                end tell
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Added text to slide {slide_index} in {presentation_name}")
            
            return {
                "status": "success",
                "presentation_name": presentation_name,
                "slide_index": slide_index,
                "text_length": len(text),
                "placeholder": placeholder,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to add text via AppleScript: {e}")
            raise
    
    async def save_powerpoint_presentation(
        self,
        presentation_name: str,
        file_path: str,
        format: str = "pptx"
    ) -> Dict[str, Any]:
        """Save a PowerPoint presentation via AppleScript.
        
        Args:
            presentation_name: Name of the presentation
            file_path: Path to save the file
            format: File format
            
        Returns:
            Dict with operation status
        """
        try:
            # Map format to PowerPoint save format
            format_map = {
                "pptx": "save as PowerPoint presentation",
                "pdf": "save as PDF",
                "ppt": "save as PowerPoint 97-2004 presentation"
            }
            
            save_format = format_map.get(format.lower(), "save as PowerPoint presentation")
            
            script = f'''
            tell application "{self.powerpoint_app}"
                tell presentation "{presentation_name}"
                    {save_format} in POSIX file "{file_path}"
                end tell
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Saved presentation {presentation_name} to {file_path}")
            
            return {
                "status": "success",
                "presentation_name": presentation_name,
                "file_path": file_path,
                "format": format,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to save presentation via AppleScript: {e}")
            raise
    
    async def create_word_document(self, title: str = "New Document") -> Dict[str, Any]:
        """Create a new Word document via AppleScript.
        
        Args:
            title: Document title
            
        Returns:
            Dict with document information
        """
        try:
            await self.launch_word()
            
            script = f'''
            tell application "{self.word_app}"
                set newDoc to make new document
                tell newDoc
                    set content to "{title}\\n\\n"
                end tell
                return name of newDoc
            end tell
            '''
            
            result = await self.execute_applescript(script)
            logger.info(f"Created Word document via AppleScript: {title}")
            
            return {
                "status": "success",
                "title": title,
                "applescript_name": result,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to create Word document via AppleScript: {e}")
            raise
    
    async def add_text_to_document(
        self,
        document_name: str,
        text: str,
        style: str = "Normal"
    ) -> Dict[str, Any]:
        """Add text to a Word document via AppleScript.
        
        Args:
            document_name: Name of the document
            text: Text to add
            style: Text style
            
        Returns:
            Dict with operation status
        """
        try:
            # Escape quotes in text
            escaped_text = text.replace('"', '\\"')
            
            script = f'''
            tell application "{self.word_app}"
                tell document "{document_name}"
                    set content to (content & "{escaped_text}\\n")
                end tell
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Added text to document {document_name}")
            
            return {
                "status": "success",
                "document_name": document_name,
                "text_length": len(text),
                "style": style,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to add text to document via AppleScript: {e}")
            raise
    
    async def save_word_document(
        self,
        document_name: str,
        file_path: str,
        format: str = "docx"
    ) -> Dict[str, Any]:
        """Save a Word document via AppleScript.
        
        Args:
            document_name: Name of the document
            file_path: Path to save the file
            format: File format
            
        Returns:
            Dict with operation status
        """
        try:
            # Map format to Word save format
            format_map = {
                "docx": "save as Word document",
                "pdf": "save as PDF",
                "doc": "save as Word 97-2004 document",
                "rtf": "save as rich text format",
                "txt": "save as plain text"
            }
            
            save_format = format_map.get(format.lower(), "save as Word document")
            
            script = f'''
            tell application "{self.word_app}"
                tell document "{document_name}"
                    {save_format} in POSIX file "{file_path}"
                end tell
            end tell
            '''
            
            await self.execute_applescript(script)
            logger.info(f"Saved document {document_name} to {file_path}")
            
            return {
                "status": "success",
                "document_name": document_name,
                "file_path": file_path,
                "format": format,
                "method": "applescript"
            }
            
        except Exception as e:
            logger.error(f"Failed to save document via AppleScript: {e}")
            raise
    
    async def get_office_version_info(self) -> Dict[str, Any]:
        """Get version information for Office applications.
        
        Returns:
            Dict with version information
        """
        try:
            powerpoint_version = "unknown"
            word_version = "unknown"
            
            # Get PowerPoint version
            try:
                script = f'''
                tell application "{self.powerpoint_app}"
                    return version
                end tell
                '''
                powerpoint_version = await self.execute_applescript(script)
            except Exception:
                pass
            
            # Get Word version
            try:
                script = f'''
                tell application "{self.word_app}"
                    return version
                end tell
                '''
                word_version = await self.execute_applescript(script)
            except Exception:
                pass
            
            return {
                "powerpoint_version": powerpoint_version,
                "word_version": word_version,
                "applescript_available": True
            }
            
        except Exception as e:
            logger.error(f"Failed to get Office version info: {e}")
            return {
                "powerpoint_version": "unknown",
                "word_version": "unknown",
                "applescript_available": False,
                "error": str(e)
            }
