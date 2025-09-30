#!/usr/bin/env python
"""
Test script to verify Excel Auto functionality
"""
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_operations import ExcelOperations
from workflow_engine import WorkflowEngine
from config import Config

def test_excel_operations():
    """Test basic Excel operations"""
    print("Testing Excel Operations...")
    
    excel_ops = ExcelOperations()
    
    # Create new file
    assert excel_ops.create_new_file(), "Failed to create new file"
    print("✓ Create new file")
    
    # Write to cells
    assert excel_ops.write_cell(1, 1, "Name"), "Failed to write to cell"
    assert excel_ops.write_cell(1, 2, "Age"), "Failed to write to cell"
    print("✓ Write to cells")
    
    # Read from cells
    value = excel_ops.read_cell(1, 1)
    assert value == "Name", f"Expected 'Name', got '{value}'"
    print("✓ Read from cells")
    
    # Write range
    data = [["Alice", 30], ["Bob", 25], ["Charlie", 35]]
    assert excel_ops.write_range(2, 1, data), "Failed to write range"
    print("✓ Write range")
    
    # Read range
    read_data = excel_ops.read_range(2, 1, 4, 2)
    assert len(read_data) == 3, "Failed to read correct range"
    print("✓ Read range")
    
    # Apply formatting
    assert excel_ops.apply_formatting(1, 1, bold=True, bg_color="CCCCCC"), "Failed to apply formatting"
    print("✓ Apply formatting")
    
    # Save file
    test_file = "/tmp/test_excel.xlsx"
    assert excel_ops.save_file(test_file), "Failed to save file"
    print("✓ Save file")
    
    # Open file
    assert excel_ops.open_file(test_file), "Failed to open file"
    print("✓ Open file")
    
    # Verify data persists
    value = excel_ops.read_cell(1, 1)
    assert value == "Name", "Data did not persist"
    print("✓ Data persists after save/load")
    
    excel_ops.close()
    print("✓ Excel Operations tests passed!\n")
    return True

def test_workflow_engine():
    """Test workflow engine"""
    print("Testing Workflow Engine...")
    
    engine = WorkflowEngine()
    
    # Create workflow
    assert engine.create_workflow("test_workflow", "Test workflow"), "Failed to create workflow"
    print("✓ Create workflow")
    
    # Add steps
    assert engine.add_step("test_workflow", "open_file", {"file_path": "/tmp/test_excel.xlsx"}), "Failed to add step"
    assert engine.add_step("test_workflow", "read_cell", {"row": 1, "col": 1}), "Failed to add step"
    print("✓ Add steps")
    
    # Get workflow info
    info = engine.get_workflow_info("test_workflow")
    assert info is not None, "Failed to get workflow info"
    assert len(info['steps']) == 2, "Wrong number of steps"
    print("✓ Get workflow info")
    
    # Execute workflow
    results = engine.execute_workflow("test_workflow")
    assert len(results) == 2, "Wrong number of results"
    print("✓ Execute workflow")
    
    # Save workflow
    workflow_file = "/tmp/test_workflow.json"
    assert engine.save_workflow("test_workflow", workflow_file), "Failed to save workflow"
    print("✓ Save workflow")
    
    # Load workflow
    engine2 = WorkflowEngine()
    assert engine2.load_workflow(workflow_file), "Failed to load workflow"
    assert "test_workflow" in engine2.get_workflow_list(), "Workflow not loaded"
    print("✓ Load workflow")
    
    # Delete workflow
    assert engine.delete_workflow("test_workflow"), "Failed to delete workflow"
    assert "test_workflow" not in engine.get_workflow_list(), "Workflow not deleted"
    print("✓ Delete workflow")
    
    print("✓ Workflow Engine tests passed!\n")
    return True

def test_config():
    """Test configuration management"""
    print("Testing Configuration...")
    
    config_file = "/tmp/test_config.json"
    config = Config(config_file)
    
    # Get default values
    app_name = config.get('app_name')
    assert app_name == 'Excel Auto', f"Wrong default app_name: {app_name}"
    print("✓ Get default values")
    
    # Set values
    config.set('test_key', 'test_value')
    assert config.get('test_key') == 'test_value', "Failed to set value"
    print("✓ Set values")
    
    # Save config
    assert config.save(), "Failed to save config"
    print("✓ Save config")
    
    # Load config
    config2 = Config(config_file)
    assert config2.get('test_key') == 'test_value', "Config not loaded correctly"
    print("✓ Load config")
    
    # Recent files
    config.add_recent_file("/tmp/file1.xlsx")
    config.add_recent_file("/tmp/file2.xlsx")
    recent = config.get_recent_files()
    assert len(recent) == 2, "Wrong number of recent files"
    assert recent[0] == "/tmp/file2.xlsx", "Recent files order is wrong"
    print("✓ Recent files management")
    
    print("✓ Configuration tests passed!\n")
    return True

def test_operations():
    """Test individual operations"""
    print("Testing Individual Operations...")
    
    engine = WorkflowEngine()
    
    # Create test file with data
    test_file = "/tmp/ops_test.xlsx"
    excel_ops = ExcelOperations()
    excel_ops.create_new_file()
    excel_ops.write_range(1, 1, [
        ["Product", "Price", "Quantity"],
        ["Apple", 100, 10],
        ["Banana", 50, 20],
        ["Orange", 75, 15]
    ])
    excel_ops.save_file(test_file)
    excel_ops.close()
    
    # Test sum operation
    engine.excel_ops.open_file(test_file)
    result = engine._op_sum_column({"col": 2, "start_row": 2, "end_row": 4})
    assert result['success'], "Sum operation failed"
    assert result['sum'] == 225, f"Wrong sum: {result['sum']}"
    print("✓ Sum column operation")
    
    # Test average operation
    result = engine._op_average_column({"col": 3, "start_row": 2, "end_row": 4})
    assert result['success'], "Average operation failed"
    assert abs(result['average'] - 15) < 0.01, f"Wrong average: {result['average']}"
    print("✓ Average column operation")
    
    engine.excel_ops.close()
    
    print("✓ Operations tests passed!\n")
    return True

def main():
    """Run all tests"""
    print("=" * 60)
    print("Excel Auto - Test Suite")
    print("=" * 60 + "\n")
    
    try:
        test_excel_operations()
        test_workflow_engine()
        test_config()
        test_operations()
        
        print("=" * 60)
        print("✓ ALL TESTS PASSED!")
        print("=" * 60)
        return 0
    
    except AssertionError as e:
        print(f"\n✗ TEST FAILED: {e}")
        return 1
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())
