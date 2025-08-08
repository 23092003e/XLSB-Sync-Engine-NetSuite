# XLSB-Sync-Engine-NetSuite Project Overview

## Purpose
An automated data integration tool designed to streamline the process of updating financial models stored in XLSB format across multiple subsidiaries. The tool processes Excel Binary Workbooks (XLSB) and synchronizes data with summary files containing entity information.

## Tech Stack
- **Language**: Python 3.x
- **Key Libraries**:
  - pandas: Data manipulation and analysis
  - xlwings: Excel COM interface for reading/writing Excel files
  - psutil: System resource monitoring
  - openpyxl: Additional Excel file handling
  - pythoncom: Windows COM interface management

## Architecture
- Modular design with separate concerns:
  - `excel_processor/`: Core processing modules
  - `scripts/`: Entry points and command-line interfaces
  - Configuration-driven with dataclass models
  - Supports both sequential and parallel processing modes

## Key Components
1. **BatchProcessor**: Handles bulk file processing with robust error handling
2. **ExcelProcessor**: Core XLSB file processing logic
3. **COMManager**: Excel COM interface management and cleanup
4. **SubsidiaryExtractor**: Extracts subsidiary information from filenames/content
5. **MemoryOptimizer**: Performance optimization for large files
6. **ConfigurationModels**: Type-safe configuration and result structures

## Processing Flow
1. Load summary data (entity mappings)
2. Extract subsidiary information from XLSB files
3. Match existing rows with summary data
4. Update existing entries and fill empty rows
5. Save updated XLSB files with comprehensive logging

# Suggested Commands for XLSB-Sync-Engine-NetSuite

## Development Setup
```bash
# Setup virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Main Usage
```bash
# Sequential processing (stable)
python src/scripts/process_entities.py ^
  --entity-folder "C:\path\to\entities" ^
  --summary-path "C:\path\to\summary\File tổng hợp Entities.xlsx" ^
  --mode seq

# Parallel processing (faster)
python src/scripts/process_entities.py ^
  --entity-folder "C:\path\to\entities" ^
  --summary-path "C:\path\to\summary\File tổng hợp Entities.xlsx" ^
  --mode par
```

## Performance Benchmarking
```bash
# Run performance benchmark
python optimize_performance.py ^
  --entity-folder "C:\path\to\entities" ^
  --summary-path "C:\path\to\summary\File tổng hợp Entities.xlsx"
```