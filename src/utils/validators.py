"""
Input validation utilities for Office 365 MCP Server
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

def validate_input(data: Dict[str, Any], schema: Dict[str, Any]) -> Dict[str, Any]:
    """Validate input data against a schema.
    
    Args:
        data: Input data to validate
        schema: Validation schema
        
    Returns:
        Validated and cleaned data
        
    Raises:
        ValueError: If validation fails
    """
    validated = {}
    
    for field, rules in schema.items():
        value = data.get(field)
        
        # Check required fields
        if rules.get("required", False) and value is None:
            raise ValueError(f"Required field '{field}' is missing")
        
        # Skip validation for optional None values
        if value is None:
            continue
        
        # Type validation
        expected_type = rules.get("type")
        if expected_type and not isinstance(value, expected_type):
            raise ValueError(f"Field '{field}' must be of type {expected_type.__name__}")
        
        # String validations
        if isinstance(value, str):
            # Length validation
            min_length = rules.get("min_length")
            max_length = rules.get("max_length")
            
            if min_length and len(value) < min_length:
                raise ValueError(f"Field '{field}' must be at least {min_length} characters")
            
            if max_length and len(value) > max_length:
                raise ValueError(f"Field '{field}' must be at most {max_length} characters")
            
            # Pattern validation
            pattern = rules.get("pattern")
            if pattern and not re.match(pattern, value):
                raise ValueError(f"Field '{field}' does not match required pattern")
        
        # Numeric validations
        if isinstance(value, (int, float)):
            min_value = rules.get("min_value")
            max_value = rules.get("max_value")
            
            if min_value is not None and value < min_value:
                raise ValueError(f"Field '{field}' must be at least {min_value}")
            
            if max_value is not None and value > max_value:
                raise ValueError(f"Field '{field}' must be at most {max_value}")
        
        # Choice validation
        choices = rules.get("choices")
        if choices and value not in choices:
            raise ValueError(f"Field '{field}' must be one of {choices}")
        
        validated[field] = value
    
    return validated

def validate_file_path(file_path: str, must_exist: bool = False, extensions: Optional[List[str]] = None) -> str:
    """Validate a file path.
    
    Args:
        file_path: Path to validate
        must_exist: Whether the file must exist
        extensions: Allowed file extensions
        
    Returns:
        Validated file path
        
    Raises:
        ValueError: If validation fails
    """
    if not file_path:
        raise ValueError("File path cannot be empty")
    
    path = Path(file_path)
    
    if must_exist and not path.exists():
        raise ValueError(f"File does not exist: {file_path}")
    
    if extensions:
        if path.suffix.lower() not in [ext.lower() for ext in extensions]:
            raise ValueError(f"File must have one of these extensions: {extensions}")
    
    return str(path.resolve())

def validate_presentation_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate presentation creation data.
    
    Args:
        data: Presentation data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "title": {
            "type": str,
            "required": True,
            "min_length": 1,
            "max_length": 255
        },
        "theme": {
            "type": str,
            "required": False,
            "choices": ["default", "modern", "classic", "minimal", "corporate"]
        },
        "template_path": {
            "type": str,
            "required": False
        }
    }
    
    validated = validate_input(data, schema)
    
    # Validate template path if provided
    if validated.get("template_path"):
        validated["template_path"] = validate_file_path(
            validated["template_path"],
            must_exist=True,
            extensions=[".pptx", ".potx"]
        )
    
    return validated

def validate_slide_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate slide creation data.
    
    Args:
        data: Slide data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "presentation_id": {
            "type": str,
            "required": True,
            "pattern": r"^[a-f0-9-]{36}$"  # UUID pattern
        },
        "layout": {
            "type": str,
            "required": False,
            "choices": [
                "Title Slide",
                "Title and Content",
                "Section Header",
                "Two Content",
                "Comparison",
                "Title Only",
                "Blank",
                "Content with Caption",
                "Picture with Caption"
            ]
        },
        "position": {
            "type": int,
            "required": False,
            "min_value": 0
        }
    }
    
    return validate_input(data, schema)

def validate_text_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate text addition data.
    
    Args:
        data: Text data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "slide_id": {
            "type": str,
            "required": True,
            "pattern": r"^[a-f0-9-]{36}$"  # UUID pattern
        },
        "text": {
            "type": str,
            "required": True,
            "min_length": 1,
            "max_length": 10000
        },
        "placeholder": {
            "type": str,
            "required": False,
            "choices": ["title", "content", "subtitle"]
        }
    }
    
    validated = validate_input(data, schema)
    
    # Validate formatting if provided
    formatting = data.get("formatting", {})
    if formatting:
        validated["formatting"] = validate_formatting_data(formatting)
    
    return validated

def validate_formatting_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate text formatting data.
    
    Args:
        data: Formatting data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "font_size": {
            "type": (int, float),
            "required": False,
            "min_value": 8,
            "max_value": 72
        },
        "font_name": {
            "type": str,
            "required": False,
            "max_length": 100
        },
        "bold": {
            "type": bool,
            "required": False
        },
        "italic": {
            "type": bool,
            "required": False
        },
        "color": {
            "type": str,
            "required": False,
            "pattern": r"^#[0-9A-Fa-f]{6}$"  # Hex color pattern
        },
        "alignment": {
            "type": str,
            "required": False,
            "choices": ["left", "center", "right", "justify"]
        }
    }
    
    return validate_input(data, schema)

def validate_image_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate image addition data.
    
    Args:
        data: Image data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "slide_id": {
            "type": str,
            "required": True,
            "pattern": r"^[a-f0-9-]{36}$"  # UUID pattern
        },
        "image_source": {
            "type": str,
            "required": True,
            "min_length": 1
        },
        "position": {
            "type": dict,
            "required": False
        },
        "size": {
            "type": dict,
            "required": False
        }
    }
    
    validated = validate_input(data, schema)
    
    # Validate image source
    image_source = validated["image_source"]
    if not (image_source.startswith(("http://", "https://")) or Path(image_source).exists()):
        raise ValueError("Image source must be a valid URL or existing file path")
    
    # Validate position and size
    if "position" in validated:
        pos_schema = {
            "x": {"type": (int, float), "required": True, "min_value": 0},
            "y": {"type": (int, float), "required": True, "min_value": 0}
        }
        validated["position"] = validate_input(validated["position"], pos_schema)
    
    if "size" in validated:
        size_schema = {
            "width": {"type": (int, float), "required": True, "min_value": 0.1},
            "height": {"type": (int, float), "required": True, "min_value": 0.1}
        }
        validated["size"] = validate_input(validated["size"], size_schema)
    
    return validated

def validate_document_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """Validate document creation data.
    
    Args:
        data: Document data to validate
        
    Returns:
        Validated data
    """
    schema = {
        "title": {
            "type": str,
            "required": False,
            "max_length": 255
        },
        "template_path": {
            "type": str,
            "required": False
        }
    }
    
    validated = validate_input(data, schema)
    
    # Validate template path if provided
    if validated.get("template_path"):
        validated["template_path"] = validate_file_path(
            validated["template_path"],
            must_exist=True,
            extensions=[".docx", ".dotx"]
        )
    
    return validated
