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

def compare_with_document_numbers(summary_old_path: str, summary_new_path: str, log_file_path: str = None):
    """
    NEW: Compare files using Document Number-based logic
    - Uses Document Number as primary key identifier
    - Highlights newly added Document Numbers (entire rows)
    - Maps entities using predefined keys
    """
    print("=" * 80)
    print("🔧 DOCUMENT NUMBER-BASED COMPARISON")
    print("=" * 80)
    
    try:
        # Use SummaryComparator's new Document Number method
        comparator = SummaryComparator()
        
        # Perform Document Number-based comparison
        success = comparator.process_summary_comparison_with_document_numbers(
            summary_old_path, summary_new_path, 
            generate_log=True, log_file_path=log_file_path
        )
        
        if success:
            print("\n✅ Document Number-based comparison completed successfully!")
            print("🆕 Newly added Document Numbers highlighted with light yellow background")
            print("📝 Changes in existing Document Numbers highlighted with light blue background")
            print("📄 Detailed Document Number-based log has been generated")
        else:
            print("\n❌ Document Number-based comparison failed!")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def compare_document_numbers_only(summary_old_path: str, summary_new_path: str):
    """
    NEW: Document Number comparison without log generation
    """
    print("=" * 80)
    print("🔧 DOCUMENT NUMBER COMPARISON (NO LOG)")
    print("=" * 80)
    
    try:
        # Use SummaryComparator's new Document Number method
        comparator = SummaryComparator()
        
        # Get comparison results
        comparison_results = comparator.compare_summary_files_by_document_number(summary_old_path, summary_new_path)
        
        new_document_numbers = comparison_results['new_document_numbers']
        entity_mapping = comparison_results['entity_mapping']
        
        print(f"\n🔍 Comparison Results:")
        print(f"   🆕 Newly added Document Numbers: {len(new_document_numbers)}")
        print(f"   🏢 Entities with new Document Numbers: {len(entity_mapping)}")
        
        if new_document_numbers:
            print(f"\n📋 New Document Number rows: {sorted(new_document_numbers)}")
        
        # Apply highlighting
        if new_document_numbers or comparison_results.get('changed_cells'):
            success = comparator.apply_document_number_highlighting_to_summary(summary_new_path, comparison_results)
            if success:
                print("\n✅ Document Number highlighting applied successfully!")
            else:
                print("\n❌ Failed to apply Document Number highlighting!")
        else:
            print("\n ℹ️ No new Document Numbers found - no highlighting needed")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def compare_with_enhanced_logic(summary_old_path: str, summary_new_path: str, log_file_path: str = None):
    """
    NEW: Enhanced comparison using oldest matching records
    - Uses ['Unit name', 'Tenant ID', 'Tenant'] key mapping
    - Finds oldest record by Date created within each group
    - For Document Numbers not in old file, compares against oldest matching entity
    - Light blue highlighting for changed cells
    - Yellow highlighting for entirely new records
    """
    print("=" * 80)
    print("🔧 ENHANCED COMPARISON (OLDEST MATCHING RECORDS)")
    print("=" * 80)
    
    try:
        # Use SummaryComparator's new enhanced method
        comparator = SummaryComparator()
        
        # Determine output directory for log
        output_dir = None
        if log_file_path:
            output_dir = os.path.dirname(log_file_path) or os.getcwd()
        else:
            output_dir = os.getcwd()
        
        # Perform enhanced comparison
        success = comparator.process_enhanced_summary_comparison(
            summary_old_path, summary_new_path, output_dir
        )
        
        if success:
            print("\n✅ Enhanced comparison completed successfully!")
            print("🔵 Changed cells highlighted with light blue background")
            print("🟡 Entirely new records highlighted with yellow background")
            print("📄 Detailed enhanced change log has been generated")
        else:
            print("\n❌ Enhanced comparison failed!")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def compare_enhanced_only(summary_old_path: str, summary_new_path: str):
    """
    NEW: Enhanced comparison without log generation
    """
    print("=" * 80)
    print("🔧 ENHANCED COMPARISON (NO LOG)")
    print("=" * 80)
    
    try:
        # Use SummaryComparator's enhanced method
        comparator = SummaryComparator()
        
        # Get comparison results
        comparison_results = comparator.compare_summary_files_with_oldest_match(summary_old_path, summary_new_path)
        
        changed_rows = comparison_results['changed_rows']
        new_rows = comparison_results['new_rows']
        changed_cells = comparison_results['changed_cells']
        
        print(f"\n🔍 Comparison Results:")
        print(f"   🔵 Rows with cell changes: {len(changed_rows)}")
        print(f"   🟡 Entirely new rows: {len(new_rows)}")
        print(f"   📊 Total cell changes: {sum(len(cols) for cols in changed_cells.values())}")
        
        if changed_rows:
            print(f"\n📋 Rows with changes: {sorted(changed_rows)}")
        
        if new_rows:
            print(f"\n📋 Entirely new rows: {sorted(new_rows)}")
        
        # Apply enhanced highlighting
        if changed_rows or new_rows:
            success = comparator.apply_enhanced_highlighting_to_summary(summary_new_path, comparison_results)
            if success:
                print("\n✅ Enhanced highlighting applied successfully!")
            else:
                print("\n❌ Failed to apply enhanced highlighting!")
        else:
            print("\n ℹ️ No changes found - no highlighting needed")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

def generate_output_files(summary_old_path: str, summary_new_path: str, output_file_path: str = None, 
                         use_document_number_logic: bool = True):
    """
    NEW: Generate Excel output files with new_rows and update_rows sheets
    """
    print("=" * 80)
    print("GENERATE OUTPUT FILES (NEW_ROWS & UPDATE_ROWS)")
    print("=" * 80)
    
    try:
        # Use SummaryComparator's new output generation method
        comparator = SummaryComparator()
        
        # Generate Excel output file
        output_path = comparator.generate_excel_output_files(
            summary_old_path, summary_new_path, output_file_path, use_document_number_logic
        )
        
        print(f"\nExcel output files generated successfully!")
        print(f"Output file: {output_path}")
        print("\nThe output file contains:")
        print("  - new_rows sheet: Completely new rows")
        print("  - update_rows sheet: Existing rows with changes")
        print("  - summary sheet: Statistics and metadata")
        
    except Exception as e:
        print(f"\nError: {e}")

def generate_output_files_with_highlighting(summary_old_path: str, summary_new_path: str, 
                                          output_file_path: str = None, use_document_number_logic: bool = True):
    """
    NEW: Generate Excel output files AND apply highlighting to original file
    """
    print("=" * 80)
    print("GENERATE OUTPUT FILES + HIGHLIGHTING")
    print("=" * 80)
    
    try:
        # Use SummaryComparator
        comparator = SummaryComparator()
        
        # First generate output files
        output_path = comparator.generate_excel_output_files(
            summary_old_path, summary_new_path, output_file_path, use_document_number_logic
        )
        
        print(f"Excel output files generated: {output_path}")
        
        # Apply NEW highlighting logic with 90-day rule
        print("Applying highlighting with 90-day rule...")
        print("   - YELLOW rows: New Document Number + Item combinations (within 90 days)")  
        print("   - BLUE cells: Updated cells in existing combinations")
        
        success = comparator.apply_highlighting_to_summary_with_90_day_rule(summary_old_path, summary_new_path)
        
        if success:
            print("Highlighting applied to original file successfully!")
            print(f"\nComplete process finished!")
            print(f"   Output file with entity format: {output_path}")
            print(f"   Summary file highlighted: {summary_new_path}")
        else:
            print("Output files generated but highlighting failed")
        
    except Exception as e:
        print(f"\nError: {e}")

def apply_highlighting_only(summary_old_path: str, summary_new_path: str):
    """
    NEW: Apply highlighting only to summary file (no output files)
    """
    print("=" * 80)
    print("APPLY HIGHLIGHTING TO SUMMARY FILE")
    print("=" * 80)
    
    try:
        # Use SummaryComparator
        comparator = SummaryComparator()
        
        print("Applying highlighting with 90-day rule...")
        print("   - YELLOW rows: New Document Number + Item combinations (within 90 days)")  
        print("   - BLUE cells: Updated cells in existing combinations")
        
        success = comparator.apply_highlighting_to_summary_with_90_day_rule(summary_old_path, summary_new_path)
        
        if success:
            print("Highlighting applied successfully!")
            print(f"   Summary file highlighted: {summary_new_path}")
        else:
            print("Highlighting failed!")
        
    except Exception as e:
        print(f"\nError: {e}")

if __name__ == "__main__":
    # Example usage
    print("Summary File Comparison Tool")
    print("-" * 40)
    
    if len(sys.argv) < 3:
        print("Usage: python compare_summary_files.py <summary_old_path> <summary_new_path> [mode]")
        print("Modes:")
        print("  highlight (default) - Compare and highlight changes (legacy logic)")
        print("  log                 - Generate detailed log only (legacy logic)")
        print("  both               - Compare, highlight, and generate log (legacy logic)")
        print("  document-number     - NEW: Document Number-based comparison with log")
        print("  document-only       - NEW: Document Number-based comparison without log")
        print("  enhanced            - NEW: Enhanced comparison with oldest matching records and log")
        print("  enhanced-only       - NEW: Enhanced comparison with oldest matching records without log")
        print("  output-files        - NEW: Generate Excel output with new_rows & update_rows sheets")
        print("  output-with-highlight - NEW: Generate output files + apply highlighting")
        print("  highlight-only      - NEW: Apply highlighting to summary file only (90-day rule)")
        print("Examples:")
        print("python compare_summary_files.py 'IPA PLC Annex T7.xlsx' 'IPA PLC Annex T8.xlsx'")
        print("python compare_summary_files.py 'Summary_Jan.xlsx' 'Summary_Feb.xlsx' log")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' both")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' document-number  # NEW")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' document-only    # NEW")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' enhanced        # NEW")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' enhanced-only   # NEW")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' output-files    # NEW - Creates Excel with sheets")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' output-with-highlight # NEW - Output + Highlighting")
        print("python compare_summary_files.py 'Previous.xlsx' 'Current.xlsx' highlight-only  # NEW - Highlight summary only")
        sys.exit(1)
    
    summary_old_file = sys.argv[1]
    summary_new_file = sys.argv[2]
    mode = sys.argv[3].lower() if len(sys.argv) > 3 else "highlight"
    
    # Validate files exist
    if not os.path.exists(summary_old_file):
        print(f"Error: Previous file not found: {summary_old_file}")
        sys.exit(1)
    
    if not os.path.exists(summary_new_file):
        print(f"Error: Current file not found: {summary_new_file}")
        sys.exit(1)
    
    # Validate mode
    valid_modes = ["highlight", "log", "both", "document-number", "document-only", "enhanced", "enhanced-only", "output-files", "output-with-highlight", "highlight-only"]
    if mode not in valid_modes:
        print(f"Invalid mode: {mode}. Use one of: {', '.join(valid_modes)}")
        sys.exit(1)
    
    print(f"Previous: {summary_old_file}")
    print(f"Current:  {summary_new_file}")
    print(f"Mode: {mode}")
    print()
    
    # Run based on mode
    if mode == "highlight":
        # Default: Compare and highlight only (no log) - Legacy logic
        processor = EnhancedExcelProcessor(DEFAULT_CONFIG)
        processor.compare_and_highlight_summary_files(
            summary_old_file, summary_new_file, 
            generate_log=False
        )
    elif mode == "log":
        # Generate detailed log only - Legacy logic
        generate_change_log_only(summary_old_file, summary_new_file)
    elif mode == "both":
        # Complete comparison with highlighting and logging - Legacy logic
        compare_with_log(summary_old_file, summary_new_file)
    elif mode == "document-number":
        # NEW: Document Number-based comparison with log
        compare_with_document_numbers(summary_old_file, summary_new_file)
    elif mode == "document-only":
        # NEW: Document Number-based comparison without log
        compare_document_numbers_only(summary_old_file, summary_new_file)
    elif mode == "enhanced":
        # NEW: Enhanced comparison with oldest matching records and log
        compare_with_enhanced_logic(summary_old_file, summary_new_file)
    elif mode == "enhanced-only":
        # NEW: Enhanced comparison with oldest matching records without log
        compare_enhanced_only(summary_old_file, summary_new_file)
    elif mode == "output-files":
        # NEW: Generate Excel output files with new_rows and update_rows sheets
        generate_output_files(summary_old_file, summary_new_file)
    elif mode == "output-with-highlight":
        # NEW: Generate output files and apply highlighting
        generate_output_files_with_highlighting(summary_old_file, summary_new_file)
    elif mode == "highlight-only":
        # NEW: Apply highlighting only to summary file
        apply_highlighting_only(summary_old_file, summary_new_file)