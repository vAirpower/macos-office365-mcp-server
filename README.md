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
git clone https://github.com/yourusername/macos-office365-mcp-server.git
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

The server uses the standard MCP protocol and can be integrated with any MCP-compatible client:

```bash
python src/office365_mcp_server.py
```

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

## 🔧 Available Tools

### PowerPoint Tools

| Tool | Description | Parameters |
|------|-------------|------------|
| `create_presentation` | Create new presentation | `title`, `theme`, `template_path` |
| `add_slide` | Add slide to presentation | `presentation_id`, `layout`, `position` |
| `add_text_to_slide` | Add text content to slide | `slide_id`, `text`, `placeholder`, `formatting` |
| `add_image_to_slide` | Add image to slide | `slide_id`, `image_source`, `position`, `size` |
| `add_speaker_notes` | Add speaker notes | `slide_id`, `notes` |
| `save_presentation` | Save presentation | `presentation_id`, `file_path`, `format` |
| `list_active_presentations` | List open presentations | None |

### Word Tools

| Tool | Description | Parameters |
|------|-------------|------------|
| `create_document` | Create new document | `title`, `template_path` |
| `add_heading` | Add heading to document | `document_id`, `text`, `level`, `style` |
| `add_paragraph` | Add paragraph to document | `document_id`, `text`, `style`, `formatting` |
| `add_list` | Add list to document | `document_id`, `items`, `list_type`, `style` |
| `add_table` | Add table to document | `document_id`, `rows`, `columns`, `data`, `style` |
| `save_document` | Save document | `document_id`, `file_path`, `format` |
| `list_active_documents` | List open documents | None |

### System Tools

| Tool | Description | Parameters |
|------|-------------|------------|
| `check_office_status` | Check Office app availability | None |

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

- **Issues**: [GitHub Issues](https://github.com/yourusername/macos-office365-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/macos-office365-mcp-server/discussions)
- **Documentation**: [Wiki](https://github.com/yourusername/macos-office365-mcp-server/wiki)

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
