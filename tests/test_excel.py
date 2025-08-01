#!/usr/bin/env python3
"""
Test script for Excel functionality in Office 365 MCP Server
"""

import asyncio
import sys
from pathlib import Path

# Add src to path (updated for new location)
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from controllers.excel_controller import ExcelController

async def test_excel_operations():
    """Test basic Excel operations."""
    print("Testing Excel Controller...")
    
    excel = ExcelController()
    
    try:
        # Test 1: Create workbook
        print("\n1. Creating workbook...")
        workbook = await excel.create_workbook(title="Test Workbook")
        workbook_id = workbook["workbook_id"]
        print(f"✓ Created workbook: {workbook_id}")
        
        # Test 2: Add worksheet
        print("\n2. Adding worksheet...")
        worksheet = await excel.add_worksheet(
            workbook_id=workbook_id,
            sheet_name="Sales Data"
        )
        print(f"✓ Added worksheet: {worksheet['sheet_name']}")
        
        # Test 3: Write cell
        print("\n3. Writing to cell...")
        await excel.write_cell(
            workbook_id=workbook_id,
            sheet_name="Sheet1",
            cell="A1",
            value="Product",
            formatting={"bold": True, "font_size": 14}
        )
        print("✓ Wrote to cell A1")
        
        # Test 4: Write range
        print("\n4. Writing data range...")
        data = [
            ["Product", "Q1", "Q2", "Q3", "Q4"],
            ["Laptops", 100, 150, 200, 180],
            ["Tablets", 80, 90, 110, 120],
            ["Phones", 200, 250, 300, 350]
        ]
        await excel.write_range(
            workbook_id=workbook_id,
            sheet_name="Sales Data",
            start_cell="A1",
            data=data,
            formatting={"bold": True}  # Bold headers
        )
        print("✓ Wrote data range")
        
        # Test 5: Add formula
        print("\n5. Adding formula...")
        await excel.add_formula(
            workbook_id=workbook_id,
            sheet_name="Sales Data",
            cell="F2",
            formula="=SUM(B2:E2)"
        )
        print("✓ Added SUM formula")
        
        # Test 6: Create chart
        print("\n6. Creating chart...")
        await excel.create_chart(
            workbook_id=workbook_id,
            sheet_name="Sales Data",
            chart_type="bar",
            data_range="A1:E4",
            chart_title="Quarterly Sales",
            position="G2"
        )
        print("✓ Created bar chart")
        
        # Test 7: List worksheets
        print("\n7. Listing worksheets...")
        sheets = await excel.list_worksheets(workbook_id)
        print(f"✓ Worksheets: {sheets}")
        
        # Test 8: Save workbook
        print("\n8. Saving workbook...")
        save_path = Path.home() / "Desktop" / "test_excel_output.xlsx"
        result = await excel.save_workbook(
            workbook_id=workbook_id,
            file_path=str(save_path)
        )
        print(f"✓ Saved to: {result['file_path']}")
        
        print("\n✅ All Excel tests passed!")
        print(f"\nWorkbook saved to: {save_path}")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_excel_operations())
