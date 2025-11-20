# Workflows Directory

This directory contains pre-defined workflow templates in JSON format.

## Available Workflows

### merge_example.json
병합 예제 - 여러 엑셀 파일의 데이터를 하나로 병합

### report_example.json
보고서 생성 - 데이터 읽기, 계산 및 서식 적용하는 보고서 생성

### sort_filter_example.json
정렬 및 필터 - 데이터를 정렬하고 특정 조건으로 필터링

## Using Workflows

### Load in GUI
1. Open Excel Auto GUI
2. Menu: `워크플로우` > `워크플로우 열기`
3. Select a workflow JSON file
4. Click `실행` to execute

### Load Programmatically
```python
from workflow_engine import WorkflowEngine

engine = WorkflowEngine()
engine.load_workflow('workflows/merge_example.json')
engine.execute_workflow('데이터_병합_예제')
```

## Creating Custom Workflows

### Using GUI
1. Menu: `워크플로우` > `새 워크플로우`
2. Add steps in the workflow editor
3. Menu: `워크플로우` > `워크플로우 저장`

### Programmatically
```python
from workflow_engine import WorkflowEngine

engine = WorkflowEngine()
engine.create_workflow("my_workflow", "My custom workflow")
engine.add_step("my_workflow", "open_file", {"file_path": "data.xlsx"})
engine.add_step("my_workflow", "sum_column", {"col": 2, "start_row": 2, "end_row": 10})
engine.save_workflow("my_workflow", "workflows/my_workflow.json")
```

## Workflow Format

Workflows are stored in JSON format:

```json
{
  "name": "workflow_name",
  "description": "Description of what this workflow does",
  "steps": [
    {
      "operation": "operation_name",
      "params": {
        "param1": "value1",
        "param2": "value2"
      }
    }
  ]
}
```

## Available Operations

See the main README.md for a complete list of available operations and their parameters.
