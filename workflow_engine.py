"""
Workflow Engine Module
Manages workflow scenarios and executes automation tasks.
"""
import json
from typing import Dict, List, Any, Callable, Optional
from excel_operations import ExcelOperations
import pandas as pd


class WorkflowEngine:
    """Engine to execute Excel automation workflows"""
    
    def __init__(self):
        self.excel_ops = ExcelOperations()
        self.workflows = {}
        self.current_workflow = None
        self.register_default_operations()
    
    def register_default_operations(self):
        """Register default workflow operations"""
        self.operations = {
            'open_file': self._op_open_file,
            'save_file': self._op_save_file,
            'read_cell': self._op_read_cell,
            'write_cell': self._op_write_cell,
            'read_range': self._op_read_range,
            'write_range': self._op_write_range,
            'format_cell': self._op_format_cell,
            'add_sheet': self._op_add_sheet,
            'delete_sheet': self._op_delete_sheet,
            'sum_column': self._op_sum_column,
            'average_column': self._op_average_column,
            'filter_data': self._op_filter_data,
            'sort_data': self._op_sort_data,
            'copy_data': self._op_copy_data,
            'merge_files': self._op_merge_files,
        }
    
    # Operation implementations
    def _op_open_file(self, params: Dict) -> Dict:
        """Open an Excel file"""
        result = self.excel_ops.open_file(
            params.get('file_path'),
            params.get('sheet_name')
        )
        return {'success': result, 'message': 'File opened' if result else 'Failed to open file'}
    
    def _op_save_file(self, params: Dict) -> Dict:
        """Save the Excel file"""
        result = self.excel_ops.save_file(params.get('file_path'))
        return {'success': result, 'message': 'File saved' if result else 'Failed to save file'}
    
    def _op_read_cell(self, params: Dict) -> Dict:
        """Read a cell value"""
        value = self.excel_ops.read_cell(params.get('row'), params.get('col'))
        return {'success': True, 'value': value}
    
    def _op_write_cell(self, params: Dict) -> Dict:
        """Write a value to a cell"""
        result = self.excel_ops.write_cell(
            params.get('row'), 
            params.get('col'), 
            params.get('value')
        )
        return {'success': result}
    
    def _op_read_range(self, params: Dict) -> Dict:
        """Read a range of cells"""
        data = self.excel_ops.read_range(
            params.get('start_row'),
            params.get('start_col'),
            params.get('end_row'),
            params.get('end_col')
        )
        return {'success': True, 'data': data}
    
    def _op_write_range(self, params: Dict) -> Dict:
        """Write data to a range"""
        result = self.excel_ops.write_range(
            params.get('start_row'),
            params.get('start_col'),
            params.get('data')
        )
        return {'success': result}
    
    def _op_format_cell(self, params: Dict) -> Dict:
        """Format a cell"""
        result = self.excel_ops.apply_formatting(
            params.get('row'),
            params.get('col'),
            params.get('bold', False),
            params.get('font_size', 11),
            params.get('bg_color'),
            params.get('align', 'left')
        )
        return {'success': result}
    
    def _op_add_sheet(self, params: Dict) -> Dict:
        """Add a new sheet"""
        result = self.excel_ops.add_sheet(params.get('sheet_name'))
        return {'success': result}
    
    def _op_delete_sheet(self, params: Dict) -> Dict:
        """Delete a sheet"""
        result = self.excel_ops.delete_sheet(params.get('sheet_name'))
        return {'success': result}
    
    def _op_sum_column(self, params: Dict) -> Dict:
        """Calculate sum of a column"""
        data = self.excel_ops.read_range(
            params.get('start_row'),
            params.get('col'),
            params.get('end_row'),
            params.get('col')
        )
        values = [row[0] for row in data if row[0] is not None and isinstance(row[0], (int, float))]
        total = sum(values)
        return {'success': True, 'sum': total}
    
    def _op_average_column(self, params: Dict) -> Dict:
        """Calculate average of a column"""
        data = self.excel_ops.read_range(
            params.get('start_row'),
            params.get('col'),
            params.get('end_row'),
            params.get('col')
        )
        values = [row[0] for row in data if row[0] is not None and isinstance(row[0], (int, float))]
        avg = sum(values) / len(values) if values else 0
        return {'success': True, 'average': avg}
    
    def _op_filter_data(self, params: Dict) -> Dict:
        """Filter data based on condition"""
        df = self.excel_ops.to_dataframe()
        if df is None:
            return {'success': False, 'message': 'Failed to convert to DataFrame'}
        
        column = params.get('column')
        operator = params.get('operator', '==')
        value = params.get('value')
        
        if operator == '==':
            filtered_df = df[df[column] == value]
        elif operator == '>':
            filtered_df = df[df[column] > value]
        elif operator == '<':
            filtered_df = df[df[column] < value]
        elif operator == '>=':
            filtered_df = df[df[column] >= value]
        elif operator == '<=':
            filtered_df = df[df[column] <= value]
        elif operator == '!=':
            filtered_df = df[df[column] != value]
        else:
            return {'success': False, 'message': 'Invalid operator'}
        
        return {'success': True, 'data': filtered_df.to_dict('records')}
    
    def _op_sort_data(self, params: Dict) -> Dict:
        """Sort data by column"""
        df = self.excel_ops.to_dataframe()
        if df is None:
            return {'success': False, 'message': 'Failed to convert to DataFrame'}
        
        column = params.get('column')
        ascending = params.get('ascending', True)
        sorted_df = df.sort_values(by=column, ascending=ascending)
        
        # Write back to worksheet
        self.excel_ops.worksheet.delete_rows(1, self.excel_ops.worksheet.max_row)
        self.excel_ops.from_dataframe(sorted_df)
        
        return {'success': True, 'message': 'Data sorted'}
    
    def _op_copy_data(self, params: Dict) -> Dict:
        """Copy data from one range to another"""
        data = self.excel_ops.read_range(
            params.get('src_start_row'),
            params.get('src_start_col'),
            params.get('src_end_row'),
            params.get('src_end_col')
        )
        result = self.excel_ops.write_range(
            params.get('dest_row'),
            params.get('dest_col'),
            data
        )
        return {'success': result}
    
    def _op_merge_files(self, params: Dict) -> Dict:
        """Merge multiple Excel files"""
        file_paths = params.get('file_paths', [])
        dfs = []
        
        for file_path in file_paths:
            temp_ops = ExcelOperations()
            if temp_ops.open_file(file_path):
                df = temp_ops.to_dataframe()
                if df is not None:
                    dfs.append(df)
                temp_ops.close()
        
        if not dfs:
            return {'success': False, 'message': 'No data to merge'}
        
        # Create new file if not already open
        if self.excel_ops.workbook is None:
            self.excel_ops.create_new_file()
        
        merged_df = pd.concat(dfs, ignore_index=True)
        self.excel_ops.from_dataframe(merged_df)
        
        return {'success': True, 'message': f'Merged {len(dfs)} files'}
    
    def create_workflow(self, name: str, description: str = "") -> bool:
        """Create a new workflow"""
        try:
            self.workflows[name] = {
                'name': name,
                'description': description,
                'steps': []
            }
            return True
        except Exception as e:
            print(f"Error creating workflow: {e}")
            return False
    
    def add_step(self, workflow_name: str, operation: str, params: Dict) -> bool:
        """Add a step to a workflow"""
        try:
            if workflow_name not in self.workflows:
                return False
            
            self.workflows[workflow_name]['steps'].append({
                'operation': operation,
                'params': params
            })
            return True
        except Exception as e:
            print(f"Error adding step: {e}")
            return False
    
    def execute_workflow(self, workflow_name: str, callback: Optional[Callable] = None) -> List[Dict]:
        """Execute a workflow"""
        if workflow_name not in self.workflows:
            return [{'success': False, 'message': 'Workflow not found'}]
        
        results = []
        workflow = self.workflows[workflow_name]
        
        for i, step in enumerate(workflow['steps']):
            operation = step['operation']
            params = step['params']
            
            if operation not in self.operations:
                results.append({
                    'success': False, 
                    'step': i + 1,
                    'message': f'Unknown operation: {operation}'
                })
                continue
            
            try:
                result = self.operations[operation](params)
                result['step'] = i + 1
                result['operation'] = operation
                results.append(result)
                
                if callback:
                    callback(i + 1, len(workflow['steps']), result)
                
            except Exception as e:
                results.append({
                    'success': False,
                    'step': i + 1,
                    'operation': operation,
                    'message': str(e)
                })
        
        return results
    
    def save_workflow(self, workflow_name: str, file_path: str) -> bool:
        """Save a workflow to a JSON file"""
        try:
            if workflow_name not in self.workflows:
                return False
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.workflows[workflow_name], f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving workflow: {e}")
            return False
    
    def load_workflow(self, file_path: str) -> bool:
        """Load a workflow from a JSON file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                workflow = json.load(f)
                self.workflows[workflow['name']] = workflow
            return True
        except Exception as e:
            print(f"Error loading workflow: {e}")
            return False
    
    def get_workflow_list(self) -> List[str]:
        """Get list of available workflows"""
        return list(self.workflows.keys())
    
    def get_workflow_info(self, workflow_name: str) -> Optional[Dict]:
        """Get information about a workflow"""
        return self.workflows.get(workflow_name)
    
    def delete_workflow(self, workflow_name: str) -> bool:
        """Delete a workflow"""
        if workflow_name in self.workflows:
            del self.workflows[workflow_name]
            return True
        return False
