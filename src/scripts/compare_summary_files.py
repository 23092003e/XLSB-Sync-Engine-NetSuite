# scripts/compare_summary_files.py
"""
Script to compare T7 vs T8 Summary files and highlight changes in T8
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from excel_processor.config import DEFAULT_CONFIG
from excel_processor.processor import EnhancedExcelProcessor
from excel_processor.summary_comparator import SummaryComparator

def compare_summary_files(summary_old_path: str, summary_new_path: str):
    """
    Compare previous period vs current period Summary files
    and highlight changes in current file
    
    Args:
        summary_old_path: Path to previous period Summary file
        summary_new_path: Path to current period Summary file
    """
    print("=" * 80)
    print("📊 SUMMARY FILE COMPARISON")
    print("=" * 80)
    
    try:
        # Initialize processor (optional - can also use SummaryComparator directly)
        processor = EnhancedExcelProcessor(DEFAULT_CONFIG)
        
        # Method 1: Use processor's integrated method
        success = processor.compare_and_highlight_summary_files(summary_old_path, summary_new_path)
        
        if success:
            print("\n✅ Summary file comparison completed successfully!")
            print("🎨 Changed rows have been highlighted in the current file with light blue background")
        else:
            print("\n❌ Summary file comparison failed!")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def compare_summary_files_direct(summary_old_path: str, summary_new_path: str):
    """
    Alternative method: Use SummaryComparator directly
    """
    print("=" * 80)
    print("📊 DIRECT SUMMARY FILE COMPARISON")
    print("=" * 80)
    
    try:
        # Method 2: Use SummaryComparator directly
        comparator = SummaryComparator()
        
        # Compare files and get changed rows
        comparison_results = comparator.compare_summary_files(summary_old_path, summary_new_path)
        changed_rows = comparison_results['changed_rows']
        
        print(f"\n🔍 Found {len(changed_rows)} changed/new rows:")
        for row_num in sorted(changed_rows):
            print(f"   - Row {row_num}")
        
        # Apply highlighting
        if changed_rows:
            success = comparator.apply_highlighting_to_summary(summary_new_path, changed_rows)
            if success:
                print("\n✅ Highlighting applied successfully!")
            else:
                print("\n❌ Failed to apply highlighting!")
        else:
            print("\n ℹ️ No changes found - no highlighting needed")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def generate_change_log_only(summary_old_path: str, summary_new_path: str, log_file_path: str = None):
    """
    Generate detailed change log without highlighting
    """
    print("=" * 80)
    print("📝 GENERATE CHANGE LOG ONLY")
    print("=" * 80)
    
    try:
        # Use SummaryComparator directly for log generation
        comparator = SummaryComparator()
        
        # Generate detailed log
        log_path = comparator.generate_detailed_change_log(summary_old_path, summary_new_path, log_file_path)
        
        print(f"\n✅ Change log generated successfully!")
        print(f"📄 Log file: {log_path}")
        print("\n📋 The log file contains:")
        print("  • Summary statistics")
        print("  • Column change frequency")
        print("  • Row-level details for each change")
        print("  • Before/after values for modified columns")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def compare_with_log(summary_old_path: str, summary_new_path: str, log_file_path: str = None):
    """
    Compare files, highlight changes, AND generate detailed log
    """
    print("=" * 80)
    print("📊 COMPLETE COMPARISON WITH LOG")
    print("=" * 80)
    
    try:
        # Initialize processor
        processor = EnhancedExcelProcessor(DEFAULT_CONFIG)
        
        # Perform comparison with log generation
        success = processor.compare_and_highlight_summary_files(
            summary_old_path, summary_new_path, 
            generate_log=True, log_file_path=log_file_path
        )
        
        if success:
            print("\n✅ Complete comparison with logging completed successfully!")
            print("🎨 Changed rows have been highlighted in the current file")
            print("📄 Detailed change log has been generated")
        else:
            print("\n❌ Comparison failed!")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    # Example usage
    print("📋 Summary File Comparison Tool")
    print("-" * 40)
    
    if len(sys.argv) < 3:
        print("Usage: python compare_summary_files.py <summary_old_path> <summary_new_path> [mode]")
        print("Modes:")
        print("  highlight (default) - Compare and highlight changes")
        print("  log                 - Generate detailed log only")
        print("  both               - Compare, highlight, and generate log")
        print("Examples:")
        print("python compare_summary_files.py 'IPA PLC Annex T7.xlsx' 'IPA PLC Annex T8.xlsx'")
        print("python compare_summary_files.py 'Summary_Jan.xlsx' 'Summary_Feb.xlsx' log")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' both")
        sys.exit(1)
    
    summary_old_file = sys.argv[1]
    summary_new_file = sys.argv[2]
    mode = sys.argv[3].lower() if len(sys.argv) > 3 else "highlight"
    
    # Validate files exist
    if not os.path.exists(summary_old_file):
        print(f"❌ Previous file not found: {summary_old_file}")
        sys.exit(1)
    
    if not os.path.exists(summary_new_file):
        print(f"❌ Current file not found: {summary_new_file}")
        sys.exit(1)
    
    # Validate mode
    if mode not in ["highlight", "log", "both"]:
        print(f"❌ Invalid mode: {mode}. Use 'highlight', 'log', or 'both'")
        sys.exit(1)
    
    print(f"📁 Previous: {summary_old_file}")
    print(f"📁 Current:  {summary_new_file}")
    print(f"🔧 Mode: {mode}")
    print()
    
    # Run based on mode
    if mode == "highlight":
        # Default: Compare and highlight only (no log)
        processor = EnhancedExcelProcessor(DEFAULT_CONFIG)
        processor.compare_and_highlight_summary_files(
            summary_old_file, summary_new_file, 
            generate_log=False
        )
    elif mode == "log":
        # Generate detailed log only
        generate_change_log_only(summary_old_file, summary_new_file)
    elif mode == "both":
        # Complete comparison with highlighting and logging
        compare_with_log(summary_old_file, summary_new_file)