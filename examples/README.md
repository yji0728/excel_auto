# Examples Directory

This directory contains example Excel files for testing and demonstration.

## Files

- `sample_data.xlsx` - Sample data file with product information
  - Contains: 제품명, 단가, 수량, 합계
  - Includes header formatting
  - Uses Excel formulas

## Creating Your Own Examples

You can create example files using the `usage_example.py` script:

```bash
python usage_example.py
```

This will create several example files demonstrating different features:
- Employee data
- Data analysis
- Merged data
- And more

## Using Examples in Workflows

Example files can be referenced in workflows:

```json
{
  "operation": "open_file",
  "params": {
    "file_path": "examples/sample_data.xlsx"
  }
}
```
