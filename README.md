# macOS Office 365 MCP Server

A comprehensive Model Context Protocol (MCP) server for automating Microsoft Office 365 applications (PowerPoint and Word) on macOS through AI agents like Claude and Cline.

## 🚀 Features

### PowerPoint Automation (15+ Tools)
- **Presentation Management**: Create, open, and save presentations
- **Slide Operations**: Add, modify, and delete slides with various layouts
- **Content Creation**: Insert text, images, tables, and charts
- **Formatting**: Apply themes, styles, and custom formatting
- **Speaker Notes**: Add and manage presenter notes
- **Export Options**: Save in multiple formats (PPTX, PDF, etc.)

### Word Automation (12+ Tools)
- **Document Management**: Create, open, and save documents
- **Content Creation**: Add headings, paragraphs, lists, and tables
- **Formatting**: Apply styles, fonts, and custom formatting
- **Template Support**: Use custom templates and themes
- **Export Options**: Save in multiple formats (DOCX, PDF, RTF, TXT)

### System Integration
- **AppleScript Bridge**: Native macOS Office integration
- **Python Libraries**: Cross-platform compatibility with python-pptx/docx
- **Real-time Operations**: Live interaction with running Office applications
- **Error Handling**: Comprehensive error handling and logging

## 📋 Requirements

### System Requirements
- **macOS**: 10.15 (Catalina) or later
- **Python**: 3.8 or later
- **Microsoft Office**: Office 365, Office 2019, or Office 2021 (optional for basic operations)

### Python Dependencies
- `mcp>=1.0.0` - Model Context Protocol framework
- `python-pptx>=0.6.21` - PowerPoint file manipulation
- `python-docx>=1.1.0` - Word document manipulation
- `pyobjc-framework-Cocoa>=10.0` - macOS system integration
- Additional dependencies listed in `requirements.txt`

## 🛠 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/vAirpower/macos-office365-mcp-server.git
cd macos-office365-mcp-server
```

### 2. Create Virtual Environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Test Installation
```bash
python test_basic.py
```

**What is `test_basic.py`?**

`test_basic.py` is a comprehensive validation script that verifies your installation without requiring a full MCP setup. It tests:
- ✅ **Import Validation**: Ensures all Python modules can be imported correctly
- ✅ **Logger Functionality**: Verifies logging system works properly
- ✅ **Configuration**: Tests configuration management
- ✅ **AppleScript Bridge**: Validates macOS integration components
- ✅ **Controllers**: Ensures PowerPoint and Word controllers initialize

This script helps you identify and fix installation issues before attempting to use the MCP server with AI agents.

## ⚙️ Configuration

### For Cline (VSCode Extension)

1. **Open Cline MCP Settings**:
   - In VSCode, open Command Palette (`Cmd+Shift+P`)
   - Search for "Cline: Open MCP Settings"

2. **Add Server Configuration**:
```json
{
  "mcpServers": {
    "office365-mcp-server": {
      "command": "/path/to/your/venv/bin/python3",
      "args": ["/path/to/macos-office365-mcp-server/src/office365_mcp_server.py"],
      "cwd": "/path/to/macos-office365-mcp-server"
    }
  }
}
```

3. **Update Paths**: Replace `/path/to/` with your actual installation path

### For Other MCP Clients

**Universal Compatibility**: This server works with ANY MCP-compatible client, not just Cline!

**Supported Clients:**
- ✅ **Cline (VSCode Extension)**: Full integration with detailed setup instructions above
- ✅ **Claude Desktop**: Can be configured to use this MCP server
- ✅ **Amazon Q CLI**: Compatible with MCP protocol
- ✅ **Custom MCP Clients**: Any application implementing the MCP standard
- ✅ **Command Line**: Direct execution for testing and development

**Generic Setup for Any MCP Client:**
```bash
# Direct execution
python src/office365_mcp_server.py

# With custom configuration
MCP_SERVER_CONFIG=/path/to/config.json python src/office365_mcp_server.py
```

**For Claude Desktop:**
1. Add to Claude Desktop's MCP configuration
2. Use the same command and args structure as shown in the Cline example
3. Restart Claude Desktop to load the server

**For Amazon Q CLI:**
1. Configure as an MCP server in Q's settings
2. Use standard MCP protocol commands
3. All tools will be available through Q's interface

## 🎯 Usage Examples

### PowerPoint Automation

```python
# Create a new presentation
presentation = await create_presentation(
    title="Quarterly Report",
    theme="modern"
)

# Add a slide
slide = await add_slide(
    presentation_id=presentation["presentation_id"],
    layout="Title and Content"
)

# Add content
await add_text_to_slide(
    slide_id=slide["slide_id"],
    text="Q4 Results Overview",
    placeholder="title"
)

await add_text_to_slide(
    slide_id=slide["slide_id"],
    text="• Revenue increased 15%\n• Customer satisfaction up 12%\n• New product launches successful",
    placeholder="content"
)

# Save presentation
await save_presentation(
    presentation_id=presentation["presentation_id"],
    file_path="/Users/username/Documents/Q4_Report.pptx"
)
```

### Word Automation

```python
# Create a new document
document = await create_document(title="Project Report")

# Add content
await add_heading(
    document_id=document["document_id"],
    text="Executive Summary",
    level=1
)

await add_paragraph(
    document_id=document["document_id"],
    text="This report summarizes the key findings from our Q4 analysis..."
)

await add_list(
    document_id=document["document_id"],
    items=["Increased revenue", "Improved efficiency", "Enhanced customer satisfaction"],
    list_type="bullet"
)

# Save document
await save_document(
    document_id=document["document_id"],
    file_path="/Users/username/Documents/Project_Report.docx"
)
```

## 🔧 Complete Tool Reference

### PowerPoint Tools (15 Tools)

#### **Presentation Management**

**`create_presentation`** - Create a new PowerPoint presentation
- **Parameters**: 
  - `title` (required): Presentation title
  - `theme` (optional): Theme name ("default", "modern", "classic", etc.)
  - `template_path` (optional): Path to custom template file
- **Returns**: `{"presentation_id": "uuid", "title": "...", "created_at": timestamp}`
- **Example**: Create a modern-themed quarterly report

**`save_presentation`** - Save presentation to file
- **Parameters**:
  - `presentation_id` (required): ID of presentation to save
  - `file_path` (required): Full path where to save the file
  - `format` (optional): File format ("pptx", "pdf", "png", "jpg")
- **Returns**: `{"status": "success", "file_path": "...", "format": "..."}`
- **Example**: Export presentation as PDF for distribution

**`list_active_presentations`** - List all currently open presentations
- **Parameters**: None
- **Returns**: Array of presentation metadata
- **Example**: Get overview of all open presentations for management

#### **Slide Operations**

**`add_slide`** - Add a new slide to presentation
- **Parameters**:
  - `presentation_id` (required): Target presentation ID
  - `layout` (optional): Slide layout ("Title and Content", "Title Only", "Blank", "Two Content", etc.)
  - `position` (optional): Position to insert slide (None for end)
- **Returns**: `{"slide_id": "uuid", "layout": "...", "index": number}`
- **Example**: Add content slide after title slide

#### **Content Creation**

**`add_text_to_slide`** - Add text content to a slide
- **Parameters**:
  - `slide_id` (required): Target slide ID
  - `text` (required): Text content to add
  - `placeholder` (optional): Placeholder name ("title", "content", "subtitle")
  - `formatting` (optional): Text formatting options (font, size, color, etc.)
- **Returns**: `{"status": "success", "text_length": number, "placeholder": "..."}`
- **Example**: Add bullet points to content area

**`add_image_to_slide`** - Insert image into slide
- **Parameters**:
  - `slide_id` (required): Target slide ID
  - `image_source` (required): Path to image file or URL
  - `position` (optional): Position dict with x, y coordinates
  - `size` (optional): Size dict with width, height
- **Returns**: `{"status": "success", "image_path": "...", "dimensions": {...}}`
- **Example**: Add company logo to slide header

**`add_speaker_notes`** - Add presenter notes to slide
- **Parameters**:
  - `slide_id` (required): Target slide ID
  - `notes` (required): Speaker notes content
- **Returns**: `{"status": "success", "notes_length": number}`
- **Example**: Add detailed talking points for presenter

### Word Tools (12 Tools)

#### **Document Management**

**`create_document`** - Create a new Word document
- **Parameters**:
  - `title` (optional): Document title (default: "New Document")
  - `template_path` (optional): Path to custom template
- **Returns**: `{"document_id": "uuid", "title": "...", "created_at": timestamp}`
- **Example**: Create report from company template

**`save_document`** - Save document to file
- **Parameters**:
  - `document_id` (required): ID of document to save
  - `file_path` (required): Full path where to save
  - `format` (optional): File format ("docx", "pdf", "rtf", "txt")
- **Returns**: `{"status": "success", "file_path": "...", "format": "..."}`
- **Example**: Export document as PDF for sharing

**`list_active_documents`** - List all currently open documents
- **Parameters**: None
- **Returns**: Array of document metadata
- **Example**: Get overview of all open documents

#### **Content Creation**

**`add_heading`** - Add formatted heading to document
- **Parameters**:
  - `document_id` (required): Target document ID
  - `text` (required): Heading text
  - `level` (optional): Heading level 1-6 (default: 1)
  - `style` (optional): Custom style name
- **Returns**: `{"status": "success", "heading_level": number, "text_length": number}`
- **Example**: Add "Executive Summary" as level 1 heading

**`add_paragraph`** - Add paragraph text to document
- **Parameters**:
  - `document_id` (required): Target document ID
  - `text` (required): Paragraph content
  - `style` (optional): Paragraph style name
  - `formatting` (optional): Text formatting options
- **Returns**: `{"status": "success", "text_length": number, "style": "..."}`
- **Example**: Add body text with custom formatting

**`add_list`** - Create bulleted or numbered list
- **Parameters**:
  - `document_id` (required): Target document ID
  - `items` (required): Array of list items
  - `list_type` (optional): "bullet" or "number" (default: "bullet")
  - `style` (optional): List style name
- **Returns**: `{"status": "success", "item_count": number, "list_type": "..."}`
- **Example**: Add action items as numbered list

**`add_table`** - Insert table with data
- **Parameters**:
  - `document_id` (required): Target document ID
  - `rows` (required): Number of rows
  - `columns` (required): Number of columns
  - `data` (optional): 2D array of table data
  - `style` (optional): Table style name
- **Returns**: `{"status": "success", "rows": number, "columns": number}`
- **Example**: Add financial data table with headers

### System Tools (3 Tools)

**`check_office_status`** - Verify Office application availability
- **Parameters**: None
- **Returns**: `{"powerpoint": {"available": boolean, "version": "..."}, "word": {...}}`
- **Example**: Confirm Office apps are installed and running

### Resource Tools (2 Tools)

**`get_templates`** - Get available Office templates
- **Parameters**: None
- **Returns**: List of available template files and themes
- **Example**: Browse available presentation themes

**`get_server_status`** - Get MCP server status information
- **Parameters**: None
- **Returns**: Server health, version, and capability information
- **Example**: Verify server is running correctly

### **Total: 27+ Tools Available**

**PowerPoint**: 15 tools for complete presentation automation
**Word**: 12 tools for comprehensive document creation
**System**: 3 tools for status and health monitoring
**Resources**: 2 tools for template and server management

**All tools support:**
- ✅ **Error Handling**: Comprehensive error messages and recovery
- ✅ **Validation**: Input parameter validation and sanitization
- ✅ **Logging**: Detailed operation logging for debugging
- ✅ **Cross-Platform**: Works with both native Office apps and python libraries
- ✅ **Async Operations**: Non-blocking execution for better performance

## 🔒 Security & Privacy

### What's Included
- ✅ All source code and utilities
- ✅ Configuration templates
- ✅ Documentation and examples
- ✅ Test files and validation scripts

### What's Excluded
- ❌ Personal file paths or user data
- ❌ API keys or credentials
- ❌ Private documents or templates
- ❌ System-specific configurations

### Permissions Required
- **Accessibility**: Required for AppleScript automation
- **File System**: Read/write access to specified directories
- **Application Control**: Permission to control Office applications

## 🧪 Testing

### Run Basic Tests
```bash
python test_basic.py
```

### Test with Real Office Apps
1. Open Microsoft PowerPoint or Word
2. Run the test script:
```bash
python -c "
import sys
sys.path.insert(0, 'src')
from office365_mcp_server import mcp
print('MCP Server ready for testing')
"
```

## 🐛 Troubleshooting

### Common Issues

**1. "spawn python ENOENT" Error**
- **Solution**: Update MCP configuration with full Python path
- **Fix**: Use absolute path to Python executable in virtual environment

**2. "MCP error -32000: Connection closed"**
- **Solution**: Ensure using FastMCP protocol
- **Fix**: Server now uses FastMCP for reliable connections

**3. AppleScript Permission Denied**
- **Solution**: Grant Accessibility permissions
- **Fix**: System Preferences → Security & Privacy → Accessibility

**4. Office Application Not Found**
- **Solution**: Install Microsoft Office or use python-only mode
- **Fix**: Server falls back to python-pptx/docx libraries

### Debug Mode
Enable detailed logging by setting environment variable:
```bash
export LOG_LEVEL=DEBUG
python src/office365_mcp_server.py
```

## 🤝 Contributing

1. **Fork the Repository**
2. **Create Feature Branch**: `git checkout -b feature/amazing-feature`
3. **Commit Changes**: `git commit -m 'Add amazing feature'`
4. **Push to Branch**: `git push origin feature/amazing-feature`
5. **Open Pull Request**

### Development Setup
```bash
# Install development dependencies
pip install -r requirements.txt

# Run tests
python test_basic.py

# Run linting
flake8 src/
black src/
mypy src/
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Model Context Protocol (MCP)**: Framework for AI agent integration
- **python-pptx/docx**: Cross-platform Office file manipulation
- **PyObjC**: macOS system integration
- **FastMCP**: Reliable MCP server implementation

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/vAirpower/macos-office365-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/vAirpower/macos-office365-mcp-server/discussions)
- **Documentation**: [Wiki](https://github.com/vAirpower/macos-office365-mcp-server/wiki)

## 🗺 Roadmap

### Upcoming Features
- [ ] Excel automation support
- [ ] Outlook integration
- [ ] Advanced chart creation
- [ ] Template marketplace
- [ ] Batch operations
- [ ] Cloud storage integration

### Version History
- **v1.0.0**: Initial release with PowerPoint and Word automation
- **v0.9.0**: Beta release with core functionality
- **v0.8.0**: Alpha release with basic MCP integration

---

**Made with ❤️ for the AI automation community**
