#!/usr/bin/env python3
"""
Basic test script for Office 365 MCP Server
Tests core functionality without requiring full MCP setup.
"""

import sys
import os
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_imports():
    """Test that all modules can be imported."""
    print("Testing imports...")
    
    try:
        from utils.logger import setup_logger
        print("✓ Logger import successful")
        
        from utils.config import Config
        print("✓ Config import successful")
        
        from utils.validators import validate_input
        print("✓ Validators import successful")
        
        from integrations.applescript_bridge import AppleScriptBridge
        print("✓ AppleScript bridge import successful")
        
        from controllers.powerpoint_controller import PowerPointController
        print("✓ PowerPoint controller import successful")
        
        from controllers.word_controller import WordController
        print("✓ Word controller import successful")
        
        return True
        
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False

def test_logger():
    """Test logger functionality."""
    print("\nTesting logger...")
    
    try:
        from utils.logger import setup_logger
        logger = setup_logger("test")
        logger.info("Test log message")
        print("✓ Logger working correctly")
        return True
        
    except Exception as e:
        print(f"✗ Logger test failed: {e}")
        return False

def test_config():
    """Test configuration functionality."""
    print("\nTesting configuration...")
    
    try:
        from utils.config import Config
        config = Config()
        log_level = config.get("log_level", "INFO")
        print(f"✓ Config working correctly (log_level: {log_level})")
        return True
        
    except Exception as e:
        print(f"✗ Config test failed: {e}")
        return False

def test_applescript():
    """Test AppleScript bridge (basic initialization)."""
    print("\nTesting AppleScript bridge...")
    
    try:
        # Import with absolute path to avoid relative import issues
        import sys
        import os
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
        
        from integrations.applescript_bridge import AppleScriptBridge
        bridge = AppleScriptBridge()
        print("✓ AppleScript bridge initialized successfully")
        return True
        
    except Exception as e:
        print(f"✗ AppleScript bridge test failed: {e}")
        return False

def test_controllers():
    """Test controller initialization."""
    print("\nTesting controllers...")
    
    try:
        # Import with absolute path to avoid relative import issues
        import sys
        import os
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
        
        from controllers.powerpoint_controller import PowerPointController
        from controllers.word_controller import WordController
        
        ppt_controller = PowerPointController()
        word_controller = WordController()
        
        print("✓ Controllers initialized successfully")
        return True
        
    except Exception as e:
        print(f"✗ Controller test failed: {e}")
        return False

def main():
    """Run all tests."""
    print("Office 365 MCP Server - Basic Tests")
    print("=" * 40)
    
    tests = [
        test_imports,
        test_logger,
        test_config,
        test_applescript,
        test_controllers
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
    
    print("\n" + "=" * 40)
    print(f"Tests completed: {passed}/{total} passed")
    
    if passed == total:
        print("🎉 All basic tests passed! The MCP server is ready for use.")
        return 0
    else:
        print("❌ Some tests failed. Please check the errors above.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
