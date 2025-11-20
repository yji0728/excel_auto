#!/usr/bin/env python
"""
Usage Example for Excel Auto
Demonstrates how to use the Excel automation features programmatically.
"""
from excel_operations import ExcelOperations
from workflow_engine import WorkflowEngine
from config import Config

def example_basic_operations():
    """Example 1: Basic Excel operations"""
    print("Example 1: Basic Excel Operations")
    print("-" * 50)
    
    excel_ops = ExcelOperations()
    
    # Create new file
    excel_ops.create_new_file()
    
    # Write headers
    headers = ["Name", "Age", "City", "Salary"]
    for i, header in enumerate(headers, 1):
        excel_ops.write_cell(1, i, header)
        excel_ops.apply_formatting(1, i, bold=True, bg_color="4472C4")
    
    # Write data
    data = [
        ["Alice", 30, "Seoul", 50000],
        ["Bob", 25, "Busan", 45000],
        ["Charlie", 35, "Incheon", 55000],
        ["David", 28, "Daegu", 48000],
    ]
    excel_ops.write_range(2, 1, data)
    
    # Save file
    excel_ops.save_file("/tmp/employees.xlsx")
    print("✓ Created employee file: /tmp/employees.xlsx")
    excel_ops.close()
    print()

def example_workflow():
    """Example 2: Using workflow engine"""
    print("Example 2: Workflow Engine")
    print("-" * 50)
    
    engine = WorkflowEngine()
    
    # Create a workflow for data processing
    engine.create_workflow("data_processor", "Process and format data")
    
    # Add steps
    engine.add_step("data_processor", "open_file", 
                   {"file_path": "/tmp/employees.xlsx"})
    
    engine.add_step("data_processor", "write_cell",
                   {"row": 1, "col": 5, "value": "Bonus"})
    
    engine.add_step("data_processor", "format_cell",
                   {"row": 1, "col": 5, "bold": True, "bg_color": "4472C4"})
    
    # Calculate bonuses (10% of salary)
    for row in range(2, 6):
        engine.add_step("data_processor", "write_cell",
                       {"row": row, "col": 5, "value": f"=D{row}*0.1"})
    
    engine.add_step("data_processor", "save_file",
                   {"file_path": "/tmp/employees_with_bonus.xlsx"})
    
    # Execute workflow
    print("Executing workflow...")
    results = engine.execute_workflow("data_processor")
    
    success_count = sum(1 for r in results if r.get('success'))
    print(f"✓ Workflow executed: {success_count}/{len(results)} steps successful")
    print(f"✓ Output file: /tmp/employees_with_bonus.xlsx")
    print()

def example_data_analysis():
    """Example 3: Data analysis operations"""
    print("Example 3: Data Analysis")
    print("-" * 50)
    
    engine = WorkflowEngine()
    
    # Open the file
    engine.excel_ops.open_file("/tmp/employees.xlsx")
    
    # Calculate sum of salaries
    result = engine._op_sum_column({"col": 4, "start_row": 2, "end_row": 5})
    print(f"✓ Total salary: {result['sum']}")
    
    # Calculate average salary
    result = engine._op_average_column({"col": 4, "start_row": 2, "end_row": 5})
    print(f"✓ Average salary: {result['average']}")
    
    # Write summary
    engine.excel_ops.write_cell(7, 3, "Total:")
    engine.excel_ops.write_cell(7, 4, "=SUM(D2:D5)")
    engine.excel_ops.apply_formatting(7, 3, bold=True)
    engine.excel_ops.apply_formatting(7, 4, bold=True)
    
    engine.excel_ops.write_cell(8, 3, "Average:")
    engine.excel_ops.write_cell(8, 4, "=AVERAGE(D2:D5)")
    engine.excel_ops.apply_formatting(8, 3, bold=True)
    engine.excel_ops.apply_formatting(8, 4, bold=True)
    
    engine.excel_ops.save_file("/tmp/employees_analysis.xlsx")
    print(f"✓ Analysis saved: /tmp/employees_analysis.xlsx")
    
    engine.excel_ops.close()
    print()

def example_merge_files():
    """Example 4: Merge multiple files"""
    print("Example 4: Merge Multiple Files")
    print("-" * 50)
    
    # Create sample files
    for i in range(1, 4):
        excel_ops = ExcelOperations()
        excel_ops.create_new_file()
        
        data = [
            ["Name", "Score"],
            [f"Student{i*10+1}", 85 + i],
            [f"Student{i*10+2}", 90 + i],
            [f"Student{i*10+3}", 78 + i],
        ]
        excel_ops.write_range(1, 1, data)
        excel_ops.save_file(f"/tmp/class{i}.xlsx")
        excel_ops.close()
        print(f"✓ Created /tmp/class{i}.xlsx")
    
    # Create workflow to merge
    engine = WorkflowEngine()
    engine.create_workflow("merge_classes", "Merge class data")
    
    engine.add_step("merge_classes", "merge_files",
                   {"file_paths": ["/tmp/class1.xlsx", "/tmp/class2.xlsx", "/tmp/class3.xlsx"]})
    
    engine.add_step("merge_classes", "save_file",
                   {"file_path": "/tmp/all_classes.xlsx"})
    
    # Execute
    results = engine.execute_workflow("merge_classes")
    if any(r.get('success') for r in results):
        print("✓ Files merged successfully: /tmp/all_classes.xlsx")
    print()

def example_save_load_workflow():
    """Example 5: Save and load workflows"""
    print("Example 5: Save and Load Workflows")
    print("-" * 50)
    
    engine = WorkflowEngine()
    
    # Create a reusable workflow
    engine.create_workflow("format_report", "Standard report formatting")
    
    engine.add_step("format_report", "open_file", {"file_path": "input.xlsx"})
    
    # Format header row
    for col in range(1, 6):
        engine.add_step("format_report", "format_cell",
                       {"row": 1, "col": col, "bold": True, 
                        "bg_color": "4472C4", "align": "center"})
    
    engine.add_step("format_report", "save_file", {"file_path": "output.xlsx"})
    
    # Save workflow
    workflow_file = "/tmp/format_report.json"
    engine.save_workflow("format_report", workflow_file)
    print(f"✓ Workflow saved: {workflow_file}")
    
    # Load it back
    engine2 = WorkflowEngine()
    engine2.load_workflow(workflow_file)
    print(f"✓ Workflow loaded")
    print(f"✓ Available workflows: {engine2.get_workflow_list()}")
    
    # Show workflow info
    info = engine2.get_workflow_info("format_report")
    print(f"✓ Workflow has {len(info['steps'])} steps")
    print()

def main():
    """Run all examples"""
    print("=" * 60)
    print("Excel Auto - Usage Examples")
    print("=" * 60)
    print()
    
    try:
        example_basic_operations()
        example_workflow()
        example_data_analysis()
        example_merge_files()
        example_save_load_workflow()
        
        print("=" * 60)
        print("✓ All examples completed successfully!")
        print("=" * 60)
        print()
        print("Files created:")
        print("  - /tmp/employees.xlsx")
        print("  - /tmp/employees_with_bonus.xlsx")
        print("  - /tmp/employees_analysis.xlsx")
        print("  - /tmp/class1.xlsx, class2.xlsx, class3.xlsx")
        print("  - /tmp/all_classes.xlsx")
        print("  - /tmp/format_report.json")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
