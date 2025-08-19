# excel_processor/summary_comparator.py
import pandas as pd
import xlwings as xw
from typing import Dict, List, Tuple, Set
from datetime import datetime
import os

class SummaryComparator:
    """
    Handles comparison between T7 (previous month) and T8 (current month) Summary files
    to identify and highlight changes in the Summary file.
    """
    
    def __init__(self):
        self.key_columns = ['Unit name', 'Tenant ID', 'Tenant']
        self.tracked_columns = {
            'GLA': 'GLA',
            'Start date (for model)': 'Start date (for model)', 
            'End date (for model)': 'End date (for model)',
            'Rent USD_Item (for model)': 'Rent USD_Item (for model)',
            'Rent VND_Item (for model)': 'Rent VND_Item (for model)',
            'Escalation rate (for model)': 'Escalation rate (for model)',
            'Service charge (for model)': 'Service charge (for model)',
            'Broker? (Yes/No)': 'Broker? (Yes/No)'
        }

    
    def generate_detailed_change_log(self, summary_old_path: str, summary_new_path: str, log_file_path: str = None) -> str:
        """
        Generate a detailed log file showing row-level changes with specific column modifications
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            log_file_path: Optional custom path for log file. If None, auto-generates based on input files
        
        Returns:
            Path to the generated log file
        """
        import json
        from datetime import datetime
        
        # Auto-generate log file path if not provided
        if log_file_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            old_name = os.path.splitext(os.path.basename(summary_old_path))[0]
            new_name = os.path.splitext(os.path.basename(summary_new_path))[0]
            log_file_path = f"summary_comparison_{old_name}_vs_{new_name}_{timestamp}.log"
        
        print(f"📝 Generating detailed change log: {log_file_path}")
        
        try:
            # Load both files
            old_df = self.load_summary_file(summary_old_path)
            new_df = self.load_summary_file(summary_new_path)
            
            # Create lookup for old data
            old_lookup = {}
            for idx, row in old_df.iterrows():
                key = self.normalize_row_key(row)
                if key and key != '||':
                    old_lookup[key] = row
            
            # Detailed change tracking
            detailed_changes = []
            summary_stats = {
                'total_new_rows': len(new_df),
                'total_old_rows': len(old_df),
                'changed_rows': 0,
                'new_rows': 0,
                'column_change_counts': {}
            }
            
            # Compare each row in new file
            for idx, new_row in new_df.iterrows():
                new_key = self.normalize_row_key(new_row)
                excel_row = idx + 2  # Excel row number (1-indexed + header)
                
                if not new_key or new_key == '||':
                    continue
                
                row_changes = {
                    'excel_row': excel_row,
                    'key_columns': {
                        'Unit name': str(new_row.get('Unit name', '')),
                        'Tenant ID': str(new_row.get('Tenant ID', '')),
                        'Tenant': str(new_row.get('Tenant', ''))
                    },
                    'change_type': '',
                    'modified_columns': []
                }
                
                if new_key in old_lookup:
                    old_row = old_lookup[new_key]
                    has_changes = False
                    
                    # Check each tracked column for changes
                    for new_col, old_col in self.tracked_columns.items():
                        if new_col in new_df.columns and old_col in old_df.columns:
                            new_val = self._normalize_value(new_row.get(new_col), new_col)
                            old_val = self._normalize_value(old_row.get(old_col), old_col)
                            
                            if new_val != old_val:
                                has_changes = True
                                
                                # Track raw values for better readability
                                raw_new_val = str(new_row.get(new_col, ''))
                                raw_old_val = str(old_row.get(old_col, ''))
                                
                                column_change = {
                                    'column': new_col,
                                    'old_value': raw_old_val,
                                    'new_value': raw_new_val,
                                    'normalized_old': old_val,
                                    'normalized_new': new_val
                                }
                                row_changes['modified_columns'].append(column_change)
                                
                                # Update stats
                                if new_col not in summary_stats['column_change_counts']:
                                    summary_stats['column_change_counts'][new_col] = 0
                                summary_stats['column_change_counts'][new_col] += 1
                    
                    if has_changes:
                        row_changes['change_type'] = 'MODIFIED'
                        detailed_changes.append(row_changes)
                        summary_stats['changed_rows'] += 1
                else:
                    # New row
                    row_changes['change_type'] = 'NEW'
                    # Add all column values for new rows
                    for col in self.tracked_columns.keys():
                        if col in new_df.columns:
                            raw_val = str(new_row.get(col, ''))
                            if raw_val.strip():  # Only include non-empty values
                                column_info = {
                                    'column': col,
                                    'old_value': '',
                                    'new_value': raw_val,
                                    'normalized_old': '',
                                    'normalized_new': self._normalize_value(new_row.get(col), col)
                                }
                                row_changes['modified_columns'].append(column_info)
                    
                    detailed_changes.append(row_changes)
                    summary_stats['new_rows'] += 1
            
            # Generate log content
            log_content = self._generate_log_content(
                summary_old_path, summary_new_path, detailed_changes, summary_stats
            )
            
            # Write log file
            with open(log_file_path, 'w', encoding='utf-8') as f:
                f.write(log_content)
            
            print(f"   ✅ Change log generated: {log_file_path}")
            print(f"   📊 Summary: {summary_stats['changed_rows']} modified, {summary_stats['new_rows']} new rows")
            
            return log_file_path
            
        except Exception as e:
            print(f"   ❌ Error generating change log: {e}")
            raise
    
    def _generate_log_content(self, old_path: str, new_path: str, changes: list, stats: dict) -> str:
        """Generate formatted log content"""
        from datetime import datetime
        
        content = []
        content.append("=" * 100)
        content.append("SUMMARY FILE COMPARISON - DETAILED CHANGE LOG")
        content.append("=" * 100)
        content.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        content.append(f"Previous File: {old_path}")
        content.append(f"Current File:  {new_path}")
        content.append("")
        
        # Summary statistics
        content.append("📊 SUMMARY STATISTICS")
        content.append("-" * 50)
        content.append(f"Total rows in previous file: {stats['total_old_rows']}")
        content.append(f"Total rows in current file:  {stats['total_new_rows']}")
        content.append(f"Modified rows: {stats['changed_rows']}")
        content.append(f"New rows:      {stats['new_rows']}")
        content.append(f"Total changes: {stats['changed_rows'] + stats['new_rows']}")
        content.append("")
        
        # Column change statistics
        if stats['column_change_counts']:
            content.append("📈 COLUMN CHANGE FREQUENCY")
            content.append("-" * 50)
            for col, count in sorted(stats['column_change_counts'].items(), key=lambda x: x[1], reverse=True):
                content.append(f"{col}: {count} changes")
            content.append("")
        
        # Detailed changes
        content.append("📝 DETAILED CHANGES")
        content.append("-" * 50)
        
        if not changes:
            content.append("No changes detected.")
        else:
            for i, change in enumerate(changes, 1):
                content.append(f"\n[{i}] ROW {change['excel_row']} - {change['change_type']}")
                content.append(f"    Key: Unit='{change['key_columns']['Unit name']}', "
                             f"TenantID='{change['key_columns']['Tenant ID']}', "
                             f"Tenant='{change['key_columns']['Tenant']}'")
                
                if change['modified_columns']:
                    content.append("    Modified Columns:")
                    for col_change in change['modified_columns']:
                        if change['change_type'] == 'NEW':
                            content.append(f"      • {col_change['column']}: '{col_change['new_value']}'")
                        else:
                            content.append(f"      • {col_change['column']}: '{col_change['old_value']}' → '{col_change['new_value']}'")
                content.append("")
        
        content.append("=" * 100)
        content.append("END OF CHANGE LOG")
        content.append("=" * 100)
        
        return "\n".join(content)
    
    def load_summary_file(self, file_path: str) -> pd.DataFrame:
        """Load summary file with error handling"""
        try:
            # Try different engines based on file extension
            if file_path.endswith('.xlsx'):
                engines = ['openpyxl', 'xlrd']
            elif file_path.endswith('.xls'):
                engines = ['xlrd', 'openpyxl']
            else:
                engines = ['openpyxl', 'xlrd']
            
            for engine in engines:
                try:
                    df = pd.read_excel(file_path, engine=engine)
                    print(f"   ✅ Loaded {file_path} using {engine} engine")
                    return df
                except Exception as e:
                    print(f"   ⚠️ Failed to load with {engine}: {e}")
                    continue
            
            raise Exception(f"Failed to load {file_path} with any engine")
            
        except Exception as e:
            print(f"   ❌ Error loading {file_path}: {e}")
            raise
    
    def normalize_row_key(self, row: pd.Series) -> str:
        """Create normalized key for row matching"""
        key_parts = []
        for col in self.key_columns:
            if col in row:
                value = str(row[col]).strip().lower() if pd.notna(row[col]) else ''
                key_parts.append(value)
            else:
                key_parts.append('')
        return '|'.join(key_parts)
    
    def compare_summary_files(self, summary_old_path: str, summary_new_path: str) -> Dict[str, any]:
        """
        Compare summary_old vs summary_new and return specific cells that changed
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
        
        Returns:
            Dict with 'changed_cells': dict mapping row indices to sets of changed column names,
            and 'changed_rows': set of row indices (for backward compatibility)
        """
        print(f"🔍 Comparing Summary files:")
        print(f"   Previous: {summary_old_path}")
        print(f"   Current:  {summary_new_path}")
        
        # Load both files
        old_df = self.load_summary_file(summary_old_path)
        new_df = self.load_summary_file(summary_new_path)
        
        print(f"   Previous rows: {len(old_df)}, Current rows: {len(new_df)}")
        
        # Create lookup for old data
        old_lookup = {}
        for idx, row in old_df.iterrows():
            key = self.normalize_row_key(row)
            if key and key != '||':  # Skip empty keys
                old_lookup[key] = row
        
        print(f"   Previous lookup created with {len(old_lookup)} unique keys")
        
        changed_rows = set()
        changed_cells = {}  # row_num -> set of changed column names
        
        # Compare each row in new file with corresponding row in old file
        for idx, new_row in new_df.iterrows():
            new_key = self.normalize_row_key(new_row)
            row_num = idx + 2  # +2 for Excel row number (1-indexed + header)
            
            if not new_key or new_key == '||':
                continue
                
            if new_key in old_lookup:
                old_row = old_lookup[new_key]
                
                # Check each tracked column for changes
                changed_columns = set()
                for new_col, old_col in self.tracked_columns.items():
                    if new_col in new_df.columns and old_col in old_df.columns:
                        new_val = self._normalize_value(new_row.get(new_col), new_col)
                        old_val = self._normalize_value(old_row.get(old_col), old_col)
                        
                        if new_val != old_val:
                            changed_columns.add(new_col)
                
                if changed_columns:
                    changed_rows.add(row_num)
                    changed_cells[row_num] = changed_columns
            else:
                # New row in current file - mark all tracked columns as changed
                changed_rows.add(row_num)
                changed_columns = set(self.tracked_columns.keys())
                # Only include columns that actually exist in the new dataframe
                changed_columns = {col for col in changed_columns if col in new_df.columns}
                changed_cells[row_num] = changed_columns
                print(f"   🆕 Row {row_num} is new in current file")
        
        print(f"   🎯 Found {len(changed_rows)} changed/new rows in current file")
        print(f"   🎯 Found {sum(len(cols) for cols in changed_cells.values())} individual cell changes")
        
        return {
            'changed_rows': changed_rows,
            'changed_cells': changed_cells
        }
    
    def _normalize_value(self, value, column_name: str = '') -> str:
        """Normalize value for comparison based on column type"""
        if pd.isna(value):
            return ''
        
        # Handle datetime objects
        if isinstance(value, datetime):
            return value.strftime('%m/%d/%Y')
        
        # Convert to string and normalize
        str_val = str(value).strip()
        
        # Handle date columns
        if 'date' in column_name.lower():
            try:
                # Try to parse and reformat date
                if '/' in str_val:
                    # Try mm/dd/yyyy format
                    dt = datetime.strptime(str_val, '%m/%d/%Y')
                elif '-' in str_val:
                    # Try yyyy-mm-dd format
                    dt = datetime.strptime(str_val, '%Y-%m-%d')
                else:
                    return str_val.lower()
                return dt.strftime('%m/%d/%Y')
            except:
                return str_val.lower()
        
        # Handle numeric columns (Rent, Service charge, Escalation rate)
        elif any(keyword in column_name.lower() for keyword in ['rent', 'service charge', 'escalation rate']):
            try:
                # Remove common formatting characters
                clean_val = str_val.replace(',', '').replace('$', '').replace('%', '').strip()
                if clean_val and clean_val.replace('.', '').replace('-', '').isdigit():
                    # Convert to float for consistent comparison
                    num_val = float(clean_val)
                    return str(num_val)
                else:
                    return str_val.lower()
            except:
                return str_val.lower()
        
        # Handle boolean columns (Broker)
        elif 'broker' in column_name.lower():
            normalized = str_val.lower()
            # Normalize common boolean representations
            if normalized in ['yes', 'y', 'true', '1']:
                return 'yes'
            elif normalized in ['no', 'n', 'false', '0']:
                return 'no'
            else:
                return normalized
        
        # Handle GLA (numeric)
        elif 'gla' in column_name.lower():
            try:
                clean_val = str_val.replace(',', '').strip()
                if clean_val and clean_val.replace('.', '').isdigit():
                    return str(float(clean_val))
                else:
                    return str_val.lower()
            except:
                return str_val.lower()
        
        # Default: return normalized string
        return str_val.lower()
    
    def apply_highlighting_to_summary(self, summary_path: str, comparison_result: Dict[str, any]) -> bool:
        """
        Apply highlighting to specific changed cells in the Summary file
        
        Args:
            summary_path: Path to the T8 Summary file
            comparison_result: Result from compare_summary_files containing changed_cells and changed_rows
        
        Returns:
            True if highlighting was successful, False otherwise
        """
        changed_cells = comparison_result.get('changed_cells', {})
        changed_rows = comparison_result.get('changed_rows', set())
        
        if not changed_cells and not changed_rows:
            print("   ℹ️ No cells to highlight in Summary file")
            return True
        
        try:
            print(f"🎨 Applying cell-specific highlighting to Summary file")
            print(f"   Rows with changes: {len(changed_rows)}")
            print(f"   Total changed cells: {sum(len(cols) for cols in changed_cells.values())}")
            
            # Open the file with xlwings
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.open(summary_path)
                sheet = wb.sheets[0]  # Assume first sheet
                
                # Get column mapping for the sheet
                header_row = sheet.range('1:1').value
                if not header_row:
                    print("   ❌ Could not read header row")
                    return False
                
                # Create column name to column number mapping
                col_mapping = {}
                for col_idx, col_name in enumerate(header_row):
                    if col_name:
                        col_mapping[str(col_name).strip()] = col_idx + 1
                
                highlighted_cells = 0
                for row_num, changed_column_names in changed_cells.items():
                    try:
                        for col_name in changed_column_names:
                            # Find the column number for this column name
                            col_num = None
                            
                            # Direct match first
                            if col_name in col_mapping:
                                col_num = col_mapping[col_name]
                            else:
                                # Try fuzzy matching for column names
                                for header_col, header_col_num in col_mapping.items():
                                    if col_name.lower().strip() in header_col.lower().strip() or \
                                       header_col.lower().strip() in col_name.lower().strip():
                                        col_num = header_col_num
                                        break
                            
                            if col_num:
                                # Highlight the specific cell with light blue background
                                cell = sheet.range((row_num, col_num))
                                cell.color = (173, 216, 230)  # Light blue for changed cells
                                highlighted_cells += 1
                            else:
                                print(f"   ⚠️ Could not find column '{col_name}' in header row")
                    
                    except Exception as e:
                        print(f"   ⚠️ Failed to highlight cells in row {row_num}: {e}")
                
                # Save the file
                wb.save()
                print(f"   💾 Saved Summary file with {highlighted_cells} highlighted cells")
                
                return True
                
            finally:
                wb.close()
                app.quit()
                
        except Exception as e:
            print(f"   ❌ Error applying highlighting to Summary file: {e}")
            return False
    
    def process_summary_comparison(self, summary_old_path: str, summary_new_path: str, generate_log: bool = True, log_file_path: str = None) -> bool:
        """
        Complete workflow: compare summary_old vs summary_new and highlight changes in summary_new
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            generate_log: Whether to generate detailed change log file
            log_file_path: Optional custom path for log file
        
        Returns:
            True if process completed successfully
        """
        try:
            # Validate files exist
            if not os.path.exists(summary_old_path):
                print(f"   ❌ Previous file not found: {summary_old_path}")
                return False
            
            if not os.path.exists(summary_new_path):
                print(f"   ❌ Current file not found: {summary_new_path}")
                return False
            
            # Generate detailed log if requested
            if generate_log:
                log_path = self.generate_detailed_change_log(summary_old_path, summary_new_path, log_file_path)
                print(f"   📄 Detailed log saved: {log_path}")
            
            # Compare files for highlighting (now returns detailed comparison results)
            comparison_results = self.compare_summary_files(summary_old_path, summary_new_path)
            changed_rows = comparison_results['changed_rows']
            changed_cells = comparison_results['changed_cells']
            
            # Apply cell-specific highlighting
            success = self.apply_highlighting_to_summary(summary_new_path, comparison_results)
            
            if success:
                print(f"   🎉 Summary comparison completed successfully!")
                print(f"   📊 Total rows with changes: {len(changed_rows)}")
                print(f"   📊 Total individual cells highlighted: {sum(len(cols) for cols in changed_cells.values())}")
                if generate_log:
                    print(f"   📋 Detailed change log: {log_path}")
            
            return success
            
        except Exception as e:
            print(f"   ❌ Error in summary comparison process: {e}")
            return False