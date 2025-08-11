# XLSB Sync Engine - Command Reference

## 🚀 Quick Start Commands

### Run XLSB Processing
```bash
# Standard sequential processing
python src/scripts/process_entities.py --entity-folder "path/to/xlsb/files" --summary-path "path/to/Entities.xlsx" --mode seq

# Parallel processing (recommended for better performance)
python src/scripts/process_entities.py --entity-folder "path/to/xlsb/files" --summary-path "path/to/Entities.xlsx" --mode par

# Example with actual paths
python src/scripts/process_entities.py --entity-folder "data/entities" --summary-path "data/summary/Entities.xlsx" --mode par
```

### Test Commands
```bash
# Run all tests
python -m pytest tests/ -v

# Run auto-fill logic tests only
python tests/run_auto_fill_tests.py

# Run specific test categories
python tests/run_auto_fill_tests.py --core --performance --integration

# Run with coverage analysis
python tests/run_auto_fill_tests.py --coverage
python -m pytest tests/ --cov=src --cov-report=html

# Quick smoke tests
python tests/run_auto_fill_tests.py --smoke

# Performance benchmarks only
python tests/run_auto_fill_tests.py --performance
```

### Optimization Commands
```bash
# System optimization check and test
python test_optimizations.py

# Test with specific XLSB file
python test_optimizations.py "path/to/your/file.xlsb"

# Performance benchmark comparison
python optimize_performance.py --entity-folder "path/to/xlsb/files" --summary-path "path/to/Entities.xlsx"
```

## 📋 Configuration Commands

### System Configuration
```bash
# Check system capabilities and configuration
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.config import DEFAULT_CONFIG; print(DEFAULT_CONFIG)"

# Test Excel COM interface
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.com_management import COMManager; print('COM OK' if COMManager.initialize_com() else 'COM Failed')"

# Check memory status
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.memory_optimizer import MemoryOptimizer; MemoryOptimizer.log_memory_stats()"
```

### Performance Tuning
```bash
# Full system optimization test
python test_optimizations.py

# Compare sequential vs parallel performance
python optimize_performance.py --entity-folder "your/entity/folder" --summary-path "your/summary.xlsx"

# View performance logs (if available)
type logs\performance.jsonl
```

## 🔧 Development Commands

### Code Quality
```bash
# Run linting
python -m flake8 src/ tests/

# Type checking
python -m mypy src/

# Code formatting
python -m black src/ tests/
python -m isort src/ tests/
```

### Debugging
```bash
# Run with detailed system information
python test_optimizations.py

# Test processing with specific file for debugging
python test_optimizations.py "path/to/problematic/file.xlsb"

# Check COM interface status
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.com_management import COMManager; COMManager.initialize_com(); print(COMManager.get_pool_stats()); COMManager.cleanup_com()"

# Performance profiling
python -m cProfile -o profile.stats src/scripts/process_entities.py --entity-folder "data" --summary-path "summary.xlsx" --mode seq
```

## 📊 Monitoring Commands

### Performance Monitoring
```bash
# Monitor system performance during processing
python test_optimizations.py

# Compare performance between sequential and parallel modes
python optimize_performance.py --entity-folder "your/folder" --summary-path "summary.xlsx"

# View performance logs if available
dir logs
type logs\performance.jsonl 2>nul || echo "No performance logs found"
```

### Log Analysis
```bash
# View processing logs (created in entity folder parent directory)
type "processing_log.txt"

# Check for any error messages in output
python src/scripts/process_entities.py --entity-folder "your/folder" --summary-path "summary.xlsx" --mode seq > output.log 2>&1
type output.log
```

## ⚙️ Advanced Commands

### Batch Processing
```bash
# Process multiple XLSB files in a folder (sequential mode)
python src/scripts/process_entities.py --entity-folder "folder/with/xlsb/files" --summary-path "path/to/summary.xlsx" --mode seq

# Process multiple XLSB files in parallel (faster)
python src/scripts/process_entities.py --entity-folder "folder/with/xlsb/files" --summary-path "path/to/summary.xlsx" --mode par

# Example batch processing with multiple entity folders
for /d %i in (entity_folder_1 entity_folder_2 entity_folder_3) do (
    python src/scripts/process_entities.py --entity-folder "%i" --summary-path "Entities.xlsx" --mode par
)
```

### Recovery and Maintenance
```bash
# Clean up Excel processes (Windows)
taskkill /f /im excel.exe

# Test system status before processing
python test_optimizations.py

# Check if files are accessible and not locked
python -c "import os; files=['file1.xlsb','file2.xlsb']; print([f for f in files if os.access(f, os.R_OK)])"
```

## 🎯 File Size Optimization Profiles

### Small Files (< 5MB)
```bash
# Use sequential mode for small files (less overhead)
python src/scripts/process_entities.py --entity-folder "small_files" --summary-path "summary.xlsx" --mode seq
```

### Medium Files (5-15MB)
```bash
# Use parallel mode for better performance
python src/scripts/process_entities.py --entity-folder "medium_files" --summary-path "summary.xlsx" --mode par
```

### Large Files (15-30MB) - Optimized
```bash
# Parallel mode with the optimized system (recommended)
python src/scripts/process_entities.py --entity-folder "large_files" --summary-path "summary.xlsx" --mode par

# Test optimization first
python test_optimizations.py "path/to/large/file.xlsb"
```

### Extra Large Files (30MB+) - Maximum Optimization
```bash
# Run with parallel mode and monitor performance
python optimize_performance.py --entity-folder "extra_large_files" --summary-path "summary.xlsx"

# Then process with parallel mode
python src/scripts/process_entities.py --entity-folder "extra_large_files" --summary-path "summary.xlsx" --mode par
```

## 📈 Performance Targets

### Expected Performance (after optimization):
- **Small files (< 5MB)**: 30-60 seconds
- **Medium files (5-15MB)**: 1-3 minutes  
- **Large files (15-30MB)**: 2-6 minutes
- **Extra large files (30MB+)**: 4-10 minutes

### Memory Usage:
- **Base memory**: ~200-500MB
- **Per chunk**: ~20-50MB additional
- **Peak usage**: Should not exceed 70% of system RAM

### Parallel Processing:
- **4 cores**: 4-6 workers optimal
- **8 cores**: 8-12 workers optimal  
- **16 cores**: 12-16 workers optimal
- **32+ cores**: 16-20 workers optimal

## 🆘 Troubleshooting Commands

### Common Issues
```bash
# Excel COM errors - kill all Excel processes first
taskkill /f /im excel.exe
python test_optimizations.py

# Test system configuration
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.com_management import COMManager; print('COM Test:', COMManager.initialize_com()); COMManager.cleanup_com()"

# Memory issues - check system status
python -c "import sys; sys.path.insert(0, 'src'); from excel_processor.memory_optimizer import MemoryOptimizer; MemoryOptimizer.log_memory_stats()"

# Performance issues - run benchmark
python optimize_performance.py --entity-folder "test_folder" --summary-path "test_summary.xlsx"
```

### Emergency Recovery
```bash
# Force kill all Excel processes
taskkill /f /im excel.exe

# Restart with fresh COM interface
python test_optimizations.py

# Test with small file first
python test_optimizations.py "small_test_file.xlsb"
```

---

## 📝 Notes

- **Always backup important files** before processing
- **Monitor memory usage** during large file processing
- **Use appropriate worker counts** for your system
- **Enable logging** for troubleshooting and optimization
- **Test with small files** before processing large batches

## 🔗 Quick Reference

### Most Important Commands:
```bash
# 1. Test system optimization
python test_optimizations.py

# 2. Process files (recommended)
python src/scripts/process_entities.py --entity-folder "your/folder" --summary-path "summary.xlsx" --mode par

# 3. Run performance benchmark
python optimize_performance.py --entity-folder "your/folder" --summary-path "summary.xlsx"

# 4. Test auto-fill logic
python tests/run_auto_fill_tests.py
```

### Key Files:
- **Main Processing**: `src/scripts/process_entities.py`
- **System Test**: `test_optimizations.py`
- **Performance Benchmark**: `optimize_performance.py`
- **Test Suite**: `tests/run_auto_fill_tests.py`
- **Configuration**: `src/excel_processor/config.py`
- **Documentation**: `README.md`, `OPTIMIZATION_SUMMARY.md`