# Summary File Comparison and Highlighting

This document describes the new Summary file comparison functionality that compares T7 (previous month) vs T8 (current month) files and highlights changes.

## ✅ Task Requirements Completed

### 1. Handover Column Logic ✅
- **Location**: `src/excel_processor/processor.py:1164`
- **Logic**: 
  - If UFL Status == "Handed Over" → Handover = "Y"
  - Otherwise → Handover = "N"
- **Implementation**: Already correctly implemented in `_transform_column_value` method

### 2. Entity File Highlighting ✅
- **Location**: `src/excel_processor/processor.py:1122-1124`
- **Implementation**: Highlighting is applied to Entity file when adding new rows
- **Color**: Light yellow background for added rows
- **Color**: Light red background for updated rows

### 3. Summary File Highlighting ✅
- **New File**: `src/excel_processor/summary_comparator.py`
- **Integration**: Added methods to `EnhancedExcelProcessor` class
- **Color**: Light blue background for changed rows in Summary file

## 🔍 Summary File Comparison Logic

### Key Columns for Row Matching
The following columns are used to match rows between T7 and T8:
- **Unit name**
- **Tenant ID** 
- **Tenant**

### Tracked Columns for Change Detection
Changes are detected in these columns:
- **GLA** - Compare numeric values directly
- **Start date (for model)** - Compare date values  
- **End date (for model)** - Compare date values
- **Rent USD_Item (for model)** - Compare numeric values
- **Rent VND_Item (for model)** - Compare numeric values  
- **Escalation rate (for model)** - Compare percentage or float values
- **Service charge (for model)** - Compare numeric values
- **Broker? (Yes/No)** - Compare boolean or string values

## 📝 Detailed Change Logging

The system now generates comprehensive log files showing row-level changes with specific column modifications.

### Log File Features
- **Auto-generated filename**: `summary_comparison_[old]_vs_[new]_[timestamp].log`
- **Summary statistics**: Total rows, changed rows, new rows
- **Column change frequency**: Which columns change most often
- **Row-level details**: Specific before/after values for each change
- **Change categorization**: MODIFIED vs NEW rows

### Usage Examples

```python
# Generate log only (no highlighting)
comparator = SummaryComparator()
log_path = comparator.generate_detailed_change_log(
    "Previous.xlsx", "Current.xlsx"
)

# Compare with highlighting AND log generation
processor = EnhancedExcelProcessor(DEFAULT_CONFIG)
processor.compare_and_highlight_summary_files(
    "Previous.xlsx", "Current.xlsx", 
    generate_log=True, log_file_path="custom_log.txt"
)
```

### Command Line Options

```bash
# Default: Highlight changes only
python compare_summary_files.py "Old.xlsx" "New.xlsx"

# Generate detailed log only
python compare_summary_files.py "Old.xlsx" "New.xlsx" log

# Both highlighting and logging
python compare_summary_files.py "Old.xlsx" "New.xlsx" both
```

## 🚀 Usage

### Method 1: Using Enhanced Excel Processor

```python
from excel_processor.config import DEFAULT_CONFIG
from excel_processor.processor import EnhancedExcelProcessor

# Initialize processor
processor = EnhancedExcelProcessor(DEFAULT_CONFIG)

# Compare and highlight
success = processor.compare_and_highlight_summary_files(
    summary_old_path="IPA PLC Annex T7.xlsx",
    summary_new_path="IPA PLC Annex T8.xlsx"
)
```

### Method 2: Using Summary Comparator Directly

```python
from excel_processor.summary_comparator import SummaryComparator

# Initialize comparator
comparator = SummaryComparator()

# Compare files
results = comparator.compare_summary_files(
    summary_old_path="IPA PLC Annex T7.xlsx", 
    summary_new_path="IPA PLC Annex T8.xlsx"
)

# Get changed rows
changed_rows = results['changed_rows']

# Apply highlighting
comparator.apply_highlighting_to_summary("IPA PLC Annex T8.xlsx", changed_rows)
```

### Method 3: Using the Script

```bash
python src/scripts/compare_summary_files.py "Summary_Jan.xlsx" "Summary_Feb.xlsx"
```

## 📋 Features

### 🔍 Change Detection
- **Row Matching**: Uses normalized keys from Unit name, Tenant ID, and Tenant
- **Value Comparison**: Normalizes values for consistent comparison
- **Date Handling**: Properly formats dates (MM/DD/YYYY) for comparison
- **New Row Detection**: Identifies rows that exist in T8 but not in T7

### 🎨 Highlighting
- **Color**: Light blue background (RGB: 173, 216, 230)
- **Scope**: Entire row highlighting
- **File Handling**: Opens, highlights, and saves T8 file automatically
- **Error Handling**: Robust error handling with detailed logging

### ⚡ Performance
- **Memory Efficient**: Loads files using appropriate engines
- **Error Recovery**: Falls back to different engines if initial load fails
- **Batch Processing**: Efficient row processing and highlighting

## 📊 Output

When running the comparison, you'll see output like:

```
🔍 Comparing Summary files:
   Previous: IPA PLC Annex T7.xlsx
   Current:  IPA PLC Annex T8.xlsx
   Previous rows: 150, Current rows: 155
   Previous lookup created with 450 unique keys
   📝 Row 45 changed: GLA '1500.0' → '1600.0'
   📝 Row 67 changed: Start date (for model) '01/15/2024' → '02/01/2024'
   📝 Row 89 changed: Rent USD_Item (for model) '2500.0' → '2750.0'
   📝 Row 103 changed: Broker? (Yes/No) 'no' → 'yes'
   🆕 Row 134 is new in current file
   🎯 Found 15 changed/new rows in current file

🎨 Applying highlighting to 15 rows in Summary file
   ✅ Highlighted row 45
   ✅ Highlighted row 67
   ✅ Highlighted row 89
   ✅ Highlighted row 103
   ✅ Highlighted row 134
   💾 Saved Summary file with 15 highlighted rows
   🎉 Summary comparison completed successfully!
   📊 Total changes highlighted: 15
```

## 🛠️ Technical Details

### Files Created/Modified

1. **`src/excel_processor/summary_comparator.py`** (NEW)
   - `SummaryComparator` class with comparison and highlighting logic

2. **`src/excel_processor/processor.py`** (MODIFIED)
   - Added import for `SummaryComparator`
   - Added `compare_and_highlight_summary_files()` method
   - Added `apply_summary_highlighting()` method

3. **`src/scripts/compare_summary_files.py`** (NEW)
   - Standalone script for running comparisons

4. **`SUMMARY_COMPARISON.md`** (NEW)
   - This documentation file

### Key Classes and Methods

- **`SummaryComparator`**: Main comparison class
  - `compare_summary_files()`: Compare T7 vs T8
  - `apply_highlighting_to_summary()`: Apply highlighting to T8
  - `process_summary_comparison()`: Complete workflow

- **`EnhancedExcelProcessor`**: Enhanced with comparison integration
  - `compare_and_highlight_summary_files()`: Integrated comparison method
  - `apply_summary_highlighting()`: Integrated highlighting method

## 🎯 Summary

All requirements have been successfully implemented:

✅ **Handover Logic**: Correctly maps UFL Status to Handover column  
✅ **Entity Highlighting**: Works correctly for added/updated rows  
✅ **Summary Highlighting**: New functionality for T7/T8 comparison  
✅ **Row Matching**: Uses Unit name, Tenant ID, Tenant as key columns  
✅ **Change Tracking**: Monitors GLA, Start date, End date, Total months fitout  
✅ **Separate File**: Created dedicated comparison module  

The system now provides comprehensive highlighting capabilities for both Entity and Summary files, with robust comparison logic for tracking changes between monthly Summary file versions.