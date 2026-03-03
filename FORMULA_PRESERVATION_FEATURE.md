# Formula Preservation Feature - Test Results

## Overview
The Excel Template Merger now supports the `preserveFormulas` configuration option, allowing users to protect formula cells from being overwritten during data merging.

## Feature Summary

### Configuration Option
```yaml
fieldMappings:
  - jsonPath: discountedPrice
    sheetName: Sheet1
    rowIndex: 2
    columnIndex: 2
    dataType: number
    preserveFormulas: true  # Skip cells containing formulas
```

### How It Works
- When `preserveFormulas: true`, the service checks if target cell contains a formula
- If formula is found, the cell is **skipped** and the original formula is **preserved**
- If `preserveFormulas: false` (or omitted), formulas are **overwritten** with the provided value

## Test Results

### Test Case: Formula Preservation Enabled
**Configuration:** `preserveFormulas: true`
**Input Data:**
```json
{
  "discountedPrice": 999.99
}
```

**Expected Result:** Cell C2 formula `=B2*1.1` should be preserved

**Actual Result:** ✅ SUCCESS
- Cell C2 data type: `f` (formula)
- Cell C2 value: `=B2*1.1`
- Formula was NOT overwritten

**Application Log Output:**
```
Applying mapping: jsonPath=discountedPrice, fillDirection=SINGLE, value type=Double
Skipping cell at (2,2) - contains formula and preserveFormulas is enabled
```

## Use Cases

### When to Use `preserveFormulas: true`
1. **Calculated fields** - Cell contains =SUM(), =AVERAGE(), =IF(), etc.
2. **Performance calculations** - Formulas for commission, tax, discount
3. **Dynamic values** - Cells that recalculate based on other cells
4. **Template structure** - Preserving complex spreadsheet logic

### When to Use `preserveFormulas: false` (Default)
1. **Simple data fields** - Direct values from data source
2. **Static mappings** - One-to-one data replacement
3. **Data migration** - Replacing all values regardless of content

## Implementation Details

### Code Changes

1. **ExcelTemplateConfig.java**
   - Added `Boolean preserveFormulas` field to FieldMapping class

2. **ExcelTemplateMergeService.java**
   - Added `isCellFormula()` helper method
   - Added `isPreserveFormulas()` helper method
   - Updated all fill methods to check formula status:
     - `fillSingle()`
     - `fillHorizontal()`
     - `fillVertical()`
     - `fillRange()`

3. **application.yml**
   - Updated example mappings with `preserveFormulas` configurations

### Formula Detection
Uses Apache POI `CellType.FORMULA` to detect formula cells:
```java
private boolean isCellFormula(Cell cell) {
    return cell != null && cell.getCellType() == CellType.FORMULA;
}
```

## Testing Evidence

### Files Created
- `templates/formula_template.xlsx` - Test template with formula cells
- `test_formula_preservation.sh` - Automated test script

### Test Template Structure
| Data Field | Value | Total (Formula) |
|-----------|-------|-----------------|
| Item 1    | 100   | =B2*1.1         |
| Item 2    | 200   | =B3*1.1         |
| Item 3    | 150   | =B4*1.1         |
| Summary   |       |                 |
| Total:    | =SUM(B2:B4) | |

### Verified Behavior
✅ Formulas are correctly identified
✅ Formula cells are skipped when `preserveFormulas: true`
✅ Logging clearly indicates when cells are preserved
✅ No errors or exceptions on formula cells

## Backward Compatibility
- Default value for `preserveFormulas` is `null`/`false`
- Existing configurations without this field work unchanged
- No breaking changes to existing field mappings
