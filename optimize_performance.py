#!/usr/bin/env python3
"""
Performance optimization script for XLSB processing
Analyzes and optimizes Excel processing for large files
"""

import sys
import time
import psutil
import os

# Add src directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from excel_processor.config import DEFAULT_CONFIG
from excel_processor.batch import RobustBatchProcessor

def benchmark_processing(entity_folder: str, summary_path: str):
    """Benchmark processing performance"""
    print("🚀 XLSB Performance Benchmark")
    print("=" * 50)
    
    # Get file info
    xlsb_files = [f for f in os.listdir(entity_folder) if f.endswith('.xlsb') and not f.startswith('~')]
    if not xlsb_files:
        print("❌ No XLSB files found!")
        return
    
    total_size = sum(os.path.getsize(os.path.join(entity_folder, f)) for f in xlsb_files) / (1024 * 1024)  # MB
    print(f"📁 Files to process: {len(xlsb_files)}")
    print(f"💽 Total size: {total_size:.1f} MB")
    print(f"📊 Average file size: {total_size/len(xlsb_files):.1f} MB")
    
    # System info
    memory_total = psutil.virtual_memory().total / (1024**3)
    cpu_count = psutil.cpu_count()
    print(f"🖥️  System: {memory_total:.1f}GB RAM, {cpu_count} CPU cores")
    print()
    
    # Test sequential vs parallel
    file_paths = [os.path.join(entity_folder, f) for f in xlsb_files]
    processor = RobustBatchProcessor(DEFAULT_CONFIG)
    
    print("🛡️  Testing SEQUENTIAL mode...")
    start_time = time.time()
    start_memory = psutil.Process().memory_info().rss / (1024**2)
    
    seq_results = processor.process_files_sequential_robust(file_paths, summary_path)
    
    seq_time = time.time() - start_time
    seq_memory = psutil.Process().memory_info().rss / (1024**2) - start_memory
    seq_success = sum(1 for r in seq_results if r.status == 'success')
    
    print(f"⏱️  Sequential: {seq_time:.1f}s, {seq_memory:+.1f}MB, {seq_success}/{len(file_paths)} success")
    print(f"📈 Sequential throughput: {total_size/seq_time:.2f} MB/s")
    print()
    
    # Wait and cleanup
    time.sleep(2)
    
    print("⚡ Testing PARALLEL mode...")
    start_time = time.time()
    start_memory = psutil.Process().memory_info().rss / (1024**2)
    
    par_results = processor.process_files_parallel_conservative(file_paths, summary_path)
    
    par_time = time.time() - start_time
    par_memory = psutil.Process().memory_info().rss / (1024**2) - start_memory
    par_success = sum(1 for r in par_results if r.status == 'success')
    
    print(f"⏱️  Parallel: {par_time:.1f}s, {par_memory:+.1f}MB, {par_success}/{len(file_paths)} success")
    print(f"📈 Parallel throughput: {total_size/par_time:.2f} MB/s")
    
    # Performance comparison
    speedup = seq_time / par_time if par_time > 0 else 0
    print()
    print("📊 PERFORMANCE SUMMARY")
    print("=" * 30)
    print(f"⚡ Speedup: {speedup:.2f}x faster with parallel mode")
    print(f"🎯 Recommended mode: {'Parallel' if speedup > 1.2 else 'Sequential'}")
    
    if total_size > 15:  # Large files
        print("💡 Optimizations for large files:")
        print("   • Use parallel mode with 4+ instances")
        print("   • Ensure 8GB+ RAM available")
        print("   • Close other Excel applications")
        print("   • Use SSD storage for better I/O")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Benchmark XLSB processing performance")
    parser.add_argument("--entity-folder", required=True, help="Folder containing *.xlsb files")
    parser.add_argument("--summary-path", required=True, help="Path to summary Excel file")
    
    args = parser.parse_args()
    benchmark_processing(args.entity_folder, args.summary_path)