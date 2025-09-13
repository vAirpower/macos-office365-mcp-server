[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/vairpower-macos-office365-mcp-server-badge.png)](https://mseep.ai/app/vairpower-macos-office365-mcp-server)

# macOS Office 365 MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to create and manipulate Microsoft Office documents (PowerPoint, Word, and Excel) on macOS. This is a Proof of Concept and a personal project that is not an official MCP Server.  Let me know if you have any issues.  

LinkedIn: https://www.linkedin.com/in/bluhmadam/ 

## Features

### PowerPoint
- Create presentations with custom titles and themes
- Add slides with various layouts
- Insert text, images, and speaker notes
- Save presentations in multiple formats

### Word
- Create documents from scratch or templates
- Add headings, paragraphs, lists, and tables
- Apply text formatting and styles
- Save documents in multiple formats

### Excel
- Create workbooks with multiple worksheets
- Write data to cells and ranges
- Add formulas and calculations
- Create charts (bar, line, pie)
- Apply cell formatting and styles
- Save workbooks in multiple formats

## Prerequisites

### System Requirements
- macOS 10.15 or later
- Python 3.8 or later
- Microsoft Office for Mac (PowerPoint, Word, and/or Excel)

### Python Dependencies
```bash
pip install mcp
pip install python-pptx
pip install python-docx
pip install openpyxl
pip install aiofiles
```

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/vAirpower/macos-office365-mcp-server.git
   cd macos-office365-mcp-server
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up macOS permissions**
   
   The server uses AppleScript to control Office applications. You need to grant permissions:
   
   a. Open **System Preferences** → **Security & Privacy** → **Privacy**
   b. Select **Automation** in the left sidebar
   c. Find your terminal application (Terminal, iTerm2, VS Code, etc.)
   d. Check the boxes for:
      - Microsoft PowerPoint
      - Microsoft Word
      - Microsoft Excel
   
   If prompted when first running the server, click "OK" to allow automation.

## Configuration

### MCP Client Configuration

Add the server to your MCP client configuration:

#### For Claude Desktop
Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office365": {
      "command": "python",
      "args": ["/path/to/macos-office365-mcp-server/src/office365_mcp_server.py"]
    }
  }
}
```

#### For Cline (VS Code)
Edit your Cline MCP settings:

```json
{
  "mcpServers": {
    "office365": {
      "command": "python",
      "args": ["/path/to/macos-office365-mcp-server/src/office365_mcp_server.py"]
    }
  }
}
```

## Usage

Once configured, the AI assistant can use the following tools:

### PowerPoint Tools
- `create_presentation` - Create a new presentation
- `add_slide` - Add a slide to a presentation
- `add_text_to_slide` - Add text content to a slide
- `add_image_to_slide` - Add an image to a slide
- `add_speaker_notes` - Add speaker notes to a slide
- `save_presentation` - Save the presentation to a file

### Word Tools
- `create_document` - Create a new document
- `add_heading` - Add a heading to a document
- `add_paragraph` - Add a paragraph with optional formatting
- `add_list` - Add a bulleted or numbered list
- `add_table` - Add a table with data
- `save_document` - Save the document to a file

### Excel Tools
- `create_workbook` - Create a new workbook
- `add_worksheet` - Add a worksheet to a workbook
- `write_cell` - Write data to a specific cell
- `write_range` - Write data to a range of cells
- `add_formula` - Add a formula to a cell
- `create_chart` - Create a chart from data
- `save_workbook` - Save the workbook to a file
- `list_worksheets` - List all worksheets in a workbook

### Utility Tools
- `list_active_presentations` - List all open presentations
- `list_active_documents` - List all open documents
- `list_active_workbooks` - List all open workbooks
- `check_office_status` - Check if Office apps are available

## Example Usage

### Creating a PowerPoint Presentation
```python
# AI Assistant can execute:
result = await create_presentation(
    title="Q4 Sales Report",
    theme="modern"
)

slide = await add_slide(
    presentation_id=result["presentation_id"],
    layout="Title and Content"
)

await add_text_to_slide(
    slide_id=slide["slide_id"],
    text="Revenue increased by 25%",
    placeholder="content"
)

await save_presentation(
    presentation_id=result["presentation_id"],
    file_path="~/Desktop/Q4_Sales.pptx"
)
```

### Creating a Word Document
```python
# AI Assistant can execute:
doc = await create_document(title="Project Proposal")

await add_heading(
    document_id=doc["document_id"],
    text="Executive Summary",
    level=1
)

await add_paragraph(
    document_id=doc["document_id"],
    text="This proposal outlines our approach...",
    formatting={"font_size": 12, "font_name": "Arial"}
)

await save_document(
    document_id=doc["document_id"],
    file_path="~/Desktop/proposal.docx"
)
```

### Creating an Excel Workbook
```python
# AI Assistant can execute:
workbook = await create_workbook(title="Sales Analysis")

# Add data
data = [
    ["Product", "Q1", "Q2", "Q3", "Q4"],
    ["Laptops", 100, 150, 200, 180],
    ["Tablets", 80, 90, 110, 120],
    ["Phones", 200, 250, 300, 350]
]

await write_range(
    workbook_id=workbook["workbook_id"],
    sheet_name="Sheet1",
    start_cell="A1",
    data=data,
    formatting={"bold": True}  # Bold headers
)

# Add a formula
await add_formula(
    workbook_id=workbook["workbook_id"],
    sheet_name="Sheet1",
    cell="F2",
    formula="=SUM(B2:E2)"
)

# Create a chart
await create_chart(
    workbook_id=workbook["workbook_id"],
    sheet_name="Sheet1",
    chart_type="bar",
    data_range="A1:E4",
    chart_title="Quarterly Sales",
    position="G2"
)

await save_workbook(
    workbook_id=workbook["workbook_id"],
    file_path="~/Desktop/sales_analysis.xlsx"
)
```

## Troubleshooting

### Permission Issues
If you see "Not authorized to send Apple events", ensure:
1. Your terminal has automation permissions for Office apps
2. Office applications are installed and have been opened at least once
3. You may need to restart your terminal after granting permissions

### Import Errors
If you see import errors:
1. Ensure all dependencies are installed: `pip install -r requirements.txt`
2. Check Python version: `python --version` (should be 3.8+)
3. Verify MCP is installed: `pip show mcp`

### Office Not Found
If Office applications aren't detected:
1. Ensure Microsoft Office for Mac is installed
2. Try opening PowerPoint/Word/Excel manually first
3. Check if Office is installed in the standard Applications folder

## Development

### Project Structure
```
macos-office365-mcp-server/
├── src/
│   ├── office365_mcp_server.py    # Main MCP server
│   ├── controllers/
│   │   ├── powerpoint_controller.py
│   │   ├── word_controller.py
│   │   └── excel_controller.py
│   ├── integrations/
│   │   └── applescript_bridge.py  # AppleScript automation
│   └── utils/
│       ├── config.py
│       ├── logger.py
│       └── validators.py
├── requirements.txt
└── README.md
```

### Adding New Features
1. Add new methods to the appropriate controller
2. Register new tools in `office365_mcp_server.py`
3. Update this README with usage examples

## License

MIT License - see LICENSE file for details

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## Support

For issues and questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Ensure you've followed all setup steps

## Acknowledgments

- Built with the Model Context Protocol (MCP)
- Uses python-pptx, python-docx, and openpyxl for document manipulation
- AppleScript integration for native Office control
