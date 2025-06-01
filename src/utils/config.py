"""
Configuration management for Office 365 MCP Server
"""

import os
import json
from pathlib import Path
from typing import Any, Dict, Optional
from dataclasses import dataclass

@dataclass
class ServerConfig:
    """Server configuration settings."""
    log_level: str = "INFO"
    temp_dir: str = "~/tmp/office365_mcp"
    max_presentations: int = 10
    max_documents: int = 10
    enable_applescript: bool = True
    enable_cloud_api: bool = False
    
class Config:
    """Configuration manager for the MCP server."""
    
    def __init__(self, config_file: Optional[str] = None):
        self.config_file = config_file or self._get_default_config_path()
        self.settings = self._load_config()
    
    def _get_default_config_path(self) -> str:
        """Get the default configuration file path."""
        return str(Path(__file__).parent.parent.parent / "config.json")
    
    def _load_config(self) -> ServerConfig:
        """Load configuration from file or environment variables."""
        config_data = {}
        
        # Load from file if it exists
        if Path(self.config_file).exists():
            try:
                with open(self.config_file, 'r') as f:
                    config_data = json.load(f)
            except Exception:
                pass
        
        # Override with environment variables
        env_overrides = {
            "log_level": os.getenv("OFFICE365_MCP_LOG_LEVEL"),
            "temp_dir": os.getenv("OFFICE365_MCP_TEMP_DIR"),
            "max_presentations": os.getenv("OFFICE365_MCP_MAX_PRESENTATIONS"),
            "max_documents": os.getenv("OFFICE365_MCP_MAX_DOCUMENTS"),
            "enable_applescript": os.getenv("OFFICE365_MCP_ENABLE_APPLESCRIPT"),
            "enable_cloud_api": os.getenv("OFFICE365_MCP_ENABLE_CLOUD_API"),
        }
        
        # Apply non-None environment variables
        for key, value in env_overrides.items():
            if value is not None:
                if key in ["max_presentations", "max_documents"]:
                    config_data[key] = int(value)
                elif key in ["enable_applescript", "enable_cloud_api"]:
                    config_data[key] = value.lower() in ("true", "1", "yes")
                else:
                    config_data[key] = value
        
        return ServerConfig(**config_data)
    
    def save_config(self) -> None:
        """Save current configuration to file."""
        config_data = {
            "log_level": self.settings.log_level,
            "temp_dir": self.settings.temp_dir,
            "max_presentations": self.settings.max_presentations,
            "max_documents": self.settings.max_documents,
            "enable_applescript": self.settings.enable_applescript,
            "enable_cloud_api": self.settings.enable_cloud_api,
        }
        
        config_path = Path(self.config_file)
        config_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(config_path, 'w') as f:
            json.dump(config_data, f, indent=2)
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value."""
        return getattr(self.settings, key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set a configuration value."""
        if hasattr(self.settings, key):
            setattr(self.settings, key, value)
